<%@ CODEPAGE = 0 %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ��ǰ ���� ������
' History : 2007.09.11 �ѿ�� ����
'###########################################################

'0 : ANSI (�⺻��) 
'949 : �ѱ��� (EUC-KR) 
'65001 : �����ڵ� (UTF-8) 
'65535 : �����ڵ� (UTF-16)
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->
	
<%
dim oip, i,page, auction,ten,makerid,magin,auction_category
	makerid = request("makeridbox")
	ten = request("tenbox")
	auction = request("auctionbox")					'���°� �˻��� ���� ����
	magin = request("maginbox")
	auction_category = request("auction_categorybox")
	page = Request("Page") 						'������ �Ѿ�� Page ��ȣ�� ����
		if Page = "" then 							'������ �Ѿ�� Page ��ȣ�� ���ٸ�
		Page = 1 
		end if
	
set oip = new Cauctionlist        			'Ŭ���� ����
oip.FPageSize = 100							'���������� �� ��������
oip.Fcurrpage = Page
oip.frectauction = auction
oip.frectten = ten
oip.frectmakerid = makerid
oip.frectmagin = magin
oip.fauction_category = auction_category
oip.fauctionlist()								'Ŭ������ ����


Sub Drawauction(selectboxname, stats)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str ,a

	response.write "<select name='" & selectboxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if stats ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	response.write "<option value='y'"							'�ɼ��� ���� ������
		if stats ="y" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">Y</option>"
	
	response.write "<option value='n'"							'�ɼ��� ���� ������
		if stats ="n" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">N</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Drawten(selectboxname, stats)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�._selectboxname�� sub���������� ����
	dim userquery, tem_str ,a

	response.write "<select name='" & selectboxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if stats ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	response.write "<option value='y'"							'�ɼ��� ���� ������
		if stats ="y" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">Y</option>"
	
	response.write "<option value='n'"							'�ɼ��� ���� ������
		if stats ="n" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">N</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Drawmakerid(boxname, stats)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�.boxname�� sub���������� ����
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if stats ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	'����� �˻� �ɼ� ���� DB���� ��������
		userquery = "select makerid"
		userquery = userquery + " from [db_item].dbo.tbl_auction a"
		userquery = userquery + " left join [db_item].[dbo].tbl_item b"
		userquery = userquery + " on a.ten_itemid = b.itemid"
		userquery = userquery + " group by makerid"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("makerid")) then 	
				tem_str = " selected"								
			end if
			response.write "<option value='" & rsget("makerid") & "' " & tem_str & ">" & db2html(rsget("makerid")) & "</option>"
			tem_str = ""				
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
'##################################################################
Sub Drawmagin(boxname, stats)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�.boxname�� sub���������� ����
		dim userquery, tem_str ,a

	response.write "<select name='" & boxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if stats ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	response.write "<option value='20'"							'�ɼ��� ���� ������
		if stats ="20" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">20%�̻�</option>"
	
	response.write "<option value='10000'"							'�ɼ��� ���� ������
		if stats ="10000" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">20%�̸�</option>"
			
	response.write "</select>"
End Sub
'##################################################################
Sub Draw_auction_category(boxname, stats)		'�˻��ϰ����ϴ� ���� ����Ʈ �ڽ����ӿ� �ְ�, ��� �ִ� ���� �˻�.boxname�� sub���������� ����
	dim userquery, tem_str

	response.write "<select name='" & boxname & "'>"		'�˻��ϰ����ϴ� ���� ����Ʈ �������� �ϰ�
	response.write "<option value=''"							'�ɼ��� ���� ������
		if stats ="" then									'��񿡼� �˻��� ���� �����Ƿ�,
			response.write "selected"
		end if
	response.write ">����</option>"								'�����̶� �ܾ ��������.

	'����� �˻� �ɼ� ���� DB���� ��������
		userquery = "select auction_cate_code"
		userquery = userquery + " from [db_item].dbo.tbl_auction"
		userquery = userquery + " group by auction_cate_code"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("auction_cate_code")) then 	
				tem_str = " selected"								
			end if
			response.write "<option value='" & rsget("auction_cate_code") & "' " & tem_str & ">"
			if rsget("auction_cate_code") = "10010100" then
				response.write "��Ʈ/������"
			elseif rsget("auction_cate_code") = "10010200" then
				response.write "Ŭ��������"
			elseif rsget("auction_cate_code") = "10010300" then
				response.write "����Ʈ��/�޸���"
			elseif rsget("auction_cate_code") = "10010400" then
				response.write "ȭ��Ʈ/������ǰ"
			elseif rsget("auction_cate_code") = "10010500" then
				response.write "Ŭ��/����/Ȧ��"
			elseif rsget("auction_cate_code") = "10010600" then
				response.write "Į/����/��"
			elseif rsget("auction_cate_code") = "10010700" then
				response.write "�����÷�/������"
			elseif rsget("auction_cate_code") = "10010800" then
				response.write "Ǯ/������"		
			elseif rsget("auction_cate_code") = "10010900" then
				response.write "��ġ"
			elseif rsget("auction_cate_code") = "10010900" then
				response.write "������Ʈ"
			elseif rsget("auction_cate_code") = "10011000" then
				response.write "ȭ��/������ǰ"
			elseif rsget("auction_cate_code") = "10011200" then
				response.write "������ǰ��Ÿ"
				
			elseif rsget("auction_cate_code") = "10030100" then
				response.write "�����ľٹ�"
			elseif rsget("auction_cate_code") = "10030200" then
				response.write "���Ͻľٹ�"
			elseif rsget("auction_cate_code") = "10030300" then
				response.write "�ٹ���Ÿ"
			elseif rsget("auction_cate_code") = "10040100" then
				response.write "������"
			elseif rsget("auction_cate_code") = "10040200" then
				response.write "����/������/��ī"
			elseif rsget("auction_cate_code") = "10040301" then
				response.write "������"
			elseif rsget("auction_cate_code") = "10040302" then
				response.write "������"
			elseif rsget("auction_cate_code") = "10040400" then
				response.write "����/����/������"
			elseif rsget("auction_cate_code") = "10040500" then
				response.write "����/������"
			elseif rsget("auction_cate_code") = "10040600" then
				response.write "��������/Ư����"
			elseif rsget("auction_cate_code") = "10040700" then
				response.write "�ʱⱸ��Ÿ"
			
			elseif rsget("auction_cate_code") = "10050100" then
				response.write "����"
			elseif rsget("auction_cate_code") = "10050200" then
				response.write "������"
			elseif rsget("auction_cate_code") = "10050300" then
				response.write "����/���δ�"
			elseif rsget("auction_cate_code") = "10050400" then
				response.write "������/���̽�"	
			elseif rsget("auction_cate_code") = "10050500" then
				response.write "��ũ/���"
			elseif rsget("auction_cate_code") = "10050600" then
				response.write "����������"
			elseif rsget("auction_cate_code") = "10050700" then
				response.write "ĥ��/����"
			elseif rsget("auction_cate_code") = "10050900" then
				response.write "�繫�밡��"																														
			elseif rsget("auction_cate_code") = "10051000" then
				response.write "�繫��ǰ��Ÿ"
				
			elseif rsget("auction_cate_code") = "10060101" then
				response.write "�ɸ��ʹ��̾"
			elseif rsget("auction_cate_code") = "10060102" then
				response.write "�Ϸ���Ʈ���̾"
			elseif rsget("auction_cate_code") = "10060103" then
				response.write "������̾"	
			elseif rsget("auction_cate_code") = "10060104" then
				response.write "�ڵ���̵���̾"
			elseif rsget("auction_cate_code") = "10060201" then
				response.write "���͵���̾"
			elseif rsget("auction_cate_code") = "10060202" then
				response.write "������̾"
			elseif rsget("auction_cate_code") = "10060301" then
				response.write "����Ŭ�����̾"
			elseif rsget("auction_cate_code") = "10060302" then
				response.write "�ý��۴��̾"
			elseif rsget("auction_cate_code") = "99140700" then
				response.write "���̾����"
			elseif rsget("auction_cate_code") = "10060500" then
				response.write "���̾��Ÿ"
					
			elseif rsget("auction_cate_code") = "10070100" then
				response.write "����"	
			elseif rsget("auction_cate_code") = "10071000" then
				response.write "�繫����Ÿ"
			elseif rsget("auction_cate_code") = "10090200" then
				response.write "����/��������"
			elseif rsget("auction_cate_code") = "10090300" then
				response.write "������/�󺧷�"
			elseif rsget("auction_cate_code") = "10090700" then
				response.write "������/����"
			elseif rsget("auction_cate_code") = "10090800" then
				response.write "���/���/������"
			elseif rsget("auction_cate_code") = "10090900" then
				response.write "���̷���Ÿ"
				
			elseif rsget("auction_cate_code") = "99140100" then
				response.write "�̻���ǰ"
			elseif rsget("auction_cate_code") = "99140200" then
				response.write "ĳ���Ϳ�ǰ"	
			elseif rsget("auction_cate_code") = "99140300" then
				response.write "�ֹ�����/���㼱��"
			elseif rsget("auction_cate_code") = "99140400" then
				response.write "����ٹ�/�ڽ�/Ȧ��"
			elseif rsget("auction_cate_code") = "99140500" then
				response.write "Ű��Ʈ��ǰ"
			elseif rsget("auction_cate_code") = "99140600" then
				response.write "�����μ�ǰ"		
			elseif rsget("auction_cate_code") = "99140700" then
				response.write "���̵���ǰ"
			elseif rsget("auction_cate_code") = "99140800" then
				response.write "��ļ�ǰ"	
			elseif rsget("auction_cate_code") = "99140900" then
				response.write "��Ÿ��ǰ"										
			end if 
			response.write "</option>"
			tem_str = ""				
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
'##################################################################
%>

<!-- #include virtual="/admin/auction/auction.js"-->

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣��: <% drawmakerid "makeridbox" , makerid %>
			�����: <% Drawten "tenbox", ten %> 
			���ǵ��: <% Drawauction "auctionbox", auction %> 
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			����: <% Drawmagin "maginbox", magin %>
			ī�װ�: <% Draw_auction_category "auction_categorybox", auction_category %>
		</td>
	</tr>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<select name="auctionup_select" onchange="javascript:auctionup(this.value,frm);">
				<option value="">AUCTION upload ����</option>
				<option value="y">Y</option>		  
				<option value="n">N</option>	
			</select><br>					
			ī�װ����� : <div id="cd1_display" style="display:inline">	
				<select name="cd1" onchange="javascript:search1();">
					<option value="">��ī�װ�����</option>
					<option value="1">����/�繫/����</option>		  
					<option value="2">��/�ҽ�/����</option>	
				</select>	
			</div>		  
			<div id="cd2_display_1" style="display:none">	
				<select name="cd2_1" onchange="javascript:search2('cd2_1');">
					<option value="">��ī�װ�����</option>
					<option value="1">������ǰ</option>		  
					<option value="2">�ٹ�</option>	
					<option value="3">�ʱⱸ</option>
					<option value="4">�繫��ǰ</option>
					<option value="5">���̾</option>
					<option value="6">�繫���</option>
					<option value="7">���̷�</option>								
				</select>				
			</div>
			<div id="cd2_display_2" style="display:none">	
				<select name="cd2_2" onchange="javascript:search2('cd2_2');">
					<option value="">��ī�װ�����</option>
					<option value="1">������/���̵���ǰ</option>		  							
				</select>				
			</div>		  
			<div id="cd3_display_1" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10010100">��Ʈ/������</option>		  
					<option value="10010200">Ŭ��������</option>	
					<option value="10010300">����Ʈ��/�޸���</option>
					<option value="10010400">ȭ��Ʈ/������ǰ</option>
					<option value="10010500">Ŭ��/����/Ȧ��</option>
					<option value="10010600">Į/����/��</option>
					<option value="10010700">�����÷�/������</option>
					<option value="10010800">Ǯ/������</option>
					<option value="10010900">��ġ</option>				
					<option value="10011000">������Ʈ</option>												
					<option value="10011100">ȭ��/������ǰ</option>
					<option value="10011200">������ǰ��Ÿ</option>							
				</select>				
			</div>	
			<div id="cd3_display_2" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10030100">�����ľٹ�</option>		  
					<option value="10030200">���Ͻľٹ�</option>	
					<option value="10030300">�ٹ���Ÿ</option>						
				</select>				
			</div>
			<div id="cd3_display_3" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10040100">������</option>		  
					<option value="10040200">����/������/��ī</option>	
					<option value="10040301">������</option>
					<option value="10040302">������</option>
					<option value="10040400">����/����/������</option>
					<option value="10040500">����/������</option>
					<option value="10040600">��������/Ư����</option>
					<option value="10040700">�ʱⱸ��Ÿ</option>						
				</select>				
			</div>		
			<div id="cd3_display_4" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10050100">����</option>		  
					<option value="10050200">������</option>	
					<option value="10050300">����/���δ�</option>
					<option value="10050400">������/���̽�</option>
					<option value="10050500">��ũ/���</option>
					<option value="10050600">����������</option>
					<option value="10050700">ĥ��/����</option>
					<option value="10050900">�繫�밡��</option>						
					<option value="10051000">�繫��ǰ��Ÿ</option>	
				</select>	
			</div>		
			<div id="cd3_display_5" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10060101">ĳ���ʹ��̾</option>		  
					<option value="10060102">�Ϸ���Ʈ���̾</option>	
					<option value="10060103">������̾</option>
					<option value="10060104">�ڵ���̵���̾</option>
					<option value="10060201">���͵���̾</option>
					<option value="10060202">������̾</option>
					<option value="10060301">����Ŭ���÷���</option>
					<option value="10060302">�ý��۴��̾��Ÿ</option>						
					<option value="10060400">�����̸�����</option>
					<option value="10060500">���̾��Ÿ</option>
					<option value="10060100">�ҽô��̾</option>
					<option value="10060200">��ɼ����̾</option>
					<option value="10060300">�ý��۴��̾</option>													
				</select>
			</div>		
			<div id="cd3_display_6" style="display:none">	
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10070100">����</option>		  
					<option value="10071000">�繫����Ÿ</option>							
				</select>
			</div>	
			<div id="cd3_display_7" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="10090200">����/��������</option>		  
					<option value="10090300">������/�󺧷�</option>	
					<option value="10090700">������/����</option>
					<option value="10090800">���/���/������</option>
					<option value="10090900">���̷���Ÿ</option>					
				</select>
			</div>	
			<div id="cd3_display_8" style="display:none">							
				<select name="cd3" onchange="javascript:search3(this.value,frm);">
					<option value="">��ī�װ�����</option>
					<option value="99140100">�̻���ǰ</option>		  
					<option value="99140200">�ɸ��Ϳ�ǰ</option>	
					<option value="99140300">�ֹ�����/���㼱��</option>
					<option value="99140400">����ٹ�/�ڽ�/Ȧ��</option>
					<option value="99140500">Ű��Ʈ��ǰ</option>
					<option value="99140600">�����ο�ǰ</option>
					<option value="99140700">���̵���ǰ</option>
					<option value="99140800">��ļ�ǰ</option>						
					<option value="99140900">��Ÿ��ǰ</option>		
				</select>
			</div>
		</td>
		<td align="right">
			<input type="button" value="���(��ǰ)" onclick="reg('item');" class="button">
			<input type="button" value="���(�̺�Ʈ)" onclick="reg('event');" class="button">
			<input type="button" value="excel���" onclick="xmlprint(frm)" class="button">
		</td>
	</tr>
</form>	
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17">
			�˻���� : <b><%= oip.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
   		<td align="center">�̹���</td>
		<td align="center">idx(����)</td>
		<td align="center">��ǰ�ڵ�</td>
		<td align="center">��ǰ�ɼ�</td>
		<td align="center">�귣��</td>
		<td align="center">��ǰ��</td>
		<td align="center">����</td>
		<td align="center">����</td>
		<td align="center">��������</td>
		<td align="center">�����</td>
		<td align="center">ǰ��</td>
		<td align="center">����</td>
		<td align="center">���ǵ��</td>
		<td align="center">����ī�װ�</td>
		<td align="center">���</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
		<form action="/admin/auction/auction_process.asp" name="frmBuyPrc<%=i%>" method="get">			<!--for�� �ȿ��� i ���� ������ ����-->
		<input type="hidden" name="mode">	
    	<tr align="center" bgcolor="#FFFFFF">
			<td align="center"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
			<td align="center">
			<img src="<%= oip.flist(i).FImageSmall %>" width="50" height="50">
			</td>
			<td align="center"><a href="javascript:edit('<%= oip.flist(i).idx %>','<%= oip.flist(i).ten_itemid %>')"><%= oip.flist(i).idx %></a><input type="hidden" name="idx" value="<%= oip.flist(i).idx %>"></td>
			<td align="center"><%= oip.flist(i).ten_itemid %><input type="hidden" name="itemid" value="<%= oip.flist(i).ten_itemid %>"></td>
			<td align="center"><%= oip.flist(i).ten_option %></td>
			<td align="center"><%= oip.flist(i).ten_makerid %></td>
			<td align="center"><%= oip.flist(i).ten_itemname %></td>
			<td align="center"><%= oip.flist(i).fsellcash %>��</td>
			<td align="center"><%= oip.flist(i).GetCalcuMarginRate %>%</td>
			<td align="center"><%= oip.flist(i).ten_jaego %></td>
			<td align="center"><% if oip.flist(i).ten_jaego >= 10 then
					response.write "Y"
				else 
					response.write "N"
				end if %></td>
			<td align="center">
				<% if oip.flist(i).IsSoldOut then %>
					<font color=red>ǰ��</font>
    			<% end if %>
    		</td>
    		<td align="center">	
    		<% if oip.flist(i).Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oip.flist(i).Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�<br>ǰ��</font>
			<% end if %>
			</td>
			<td align="center"><%= oip.flist(i).auction_isusing %></td>
			<td align="center">
			
			<% if oip.flist(i).auction_cate_code = "10010100" then
				response.write "��Ʈ/������"
			elseif oip.flist(i).auction_cate_code = "10010200" then
				response.write "Ŭ��������"
			elseif oip.flist(i).auction_cate_code = "10010300" then
				response.write "����Ʈ��/�޸���"
			elseif oip.flist(i).auction_cate_code = "10010400" then
				response.write "ȭ��Ʈ/������ǰ"
			elseif oip.flist(i).auction_cate_code = "10010500" then
				response.write "Ŭ��/����/Ȧ��"
			elseif oip.flist(i).auction_cate_code = "10010600" then
				response.write "Į/����/��"
			elseif oip.flist(i).auction_cate_code = "10010700" then
				response.write "�����÷�/������"
			elseif oip.flist(i).auction_cate_code = "10010800" then
				response.write "Ǯ/������"		
			elseif oip.flist(i).auction_cate_code = "10010900" then
				response.write "��ġ"	
			elseif oip.flist(i).auction_cate_code = "10011000" then
				response.write "������Ʈ"
			elseif oip.flist(i).auction_cate_code = "10011100" then
				response.write "ȭ��/������ǰ"
			elseif oip.flist(i).auction_cate_code = "10011200" then
				response.write "������ǰ��Ÿ"
				
			elseif oip.flist(i).auction_cate_code = "10030100" then
				response.write "�����ľٹ�"
			elseif oip.flist(i).auction_cate_code = "10030200" then
				response.write "���Ͻľٹ�"
			elseif oip.flist(i).auction_cate_code = "10030300" then
				response.write "�ٹ���Ÿ"
			elseif oip.flist(i).auction_cate_code = "10040100" then
				response.write "������"
			elseif oip.flist(i).auction_cate_code = "10040200" then
				response.write "����/������/��ī"
			elseif oip.flist(i).auction_cate_code = "10040301" then
				response.write "������"
			elseif oip.flist(i).auction_cate_code = "10040302" then
				response.write "������"
			elseif oip.flist(i).auction_cate_code = "10040400" then
				response.write "����/����/������"
			elseif oip.flist(i).auction_cate_code = "10040500" then
				response.write "����/������"
			elseif oip.flist(i).auction_cate_code = "10040600" then
				response.write "��������/Ư����"
			elseif oip.flist(i).auction_cate_code = "10040700" then
				response.write "�ʱⱸ��Ÿ"
			
			elseif oip.flist(i).auction_cate_code = "10050100" then
				response.write "����"
			elseif oip.flist(i).auction_cate_code = "10050200" then
				response.write "������"
			elseif oip.flist(i).auction_cate_code = "10050300" then
				response.write "����/���δ�"
			elseif oip.flist(i).auction_cate_code = "10050400" then
				response.write "������/���̽�"	
			elseif oip.flist(i).auction_cate_code = "10050500" then
				response.write "��ũ/���"
			elseif oip.flist(i).auction_cate_code = "10050600" then
				response.write "����������"
			elseif oip.flist(i).auction_cate_code = "10050700" then
				response.write "ĥ��/����"
			elseif oip.flist(i).auction_cate_code = "10050900" then
				response.write "�繫�밡��"																														
			elseif oip.flist(i).auction_cate_code = "10051000" then
				response.write "�繫��ǰ��Ÿ"
				
			elseif oip.flist(i).auction_cate_code = "10060101" then
				response.write "�ɸ��ʹ��̾"
			elseif oip.flist(i).auction_cate_code = "10060102" then
				response.write "�Ϸ���Ʈ���̾"
			elseif oip.flist(i).auction_cate_code = "10060103" then
				response.write "������̾"	
			elseif oip.flist(i).auction_cate_code = "10060104" then
				response.write "�ڵ���̵���̾"
			elseif oip.flist(i).auction_cate_code = "10060201" then
				response.write "���͵���̾"
			elseif oip.flist(i).auction_cate_code = "10060202" then
				response.write "������̾"
			elseif oip.flist(i).auction_cate_code = "10060301" then
				response.write "����Ŭ�����̾"
			elseif oip.flist(i).auction_cate_code = "10060302" then
				response.write "�ý��۴��̾��Ÿ"
			elseif oip.flist(i).auction_cate_code = "99140700" then
				response.write "���̾����"
			elseif oip.flist(i).auction_cate_code = "10060500" then
				response.write "���̾��Ÿ"
			elseif oip.flist(i).auction_cate_code = "10060100" then
				response.write "�ҽô��̾"
			elseif oip.flist(i).auction_cate_code = "10060200" then
				response.write "��ɼ����̾"
			elseif oip.flist(i).auction_cate_code = "10060300" then
				response.write "�ý��۴��̾"
													
			elseif oip.flist(i).auction_cate_code = "10070100" then
				response.write "����"	
			elseif oip.flist(i).auction_cate_code = "10071000" then
				response.write "�繫����Ÿ"
			elseif oip.flist(i).auction_cate_code = "10090200" then
				response.write "����/��������"
			elseif oip.flist(i).auction_cate_code = "10090300" then
				response.write "������/�󺧷�"
			elseif oip.flist(i).auction_cate_code = "10090700" then
				response.write "������/����"
			elseif oip.flist(i).auction_cate_code = "10090800" then
				response.write "���/���/������"
			elseif oip.flist(i).auction_cate_code = "10090900" then
				response.write "���̷���Ÿ"
				
			elseif oip.flist(i).auction_cate_code = "99140100" then
				response.write "�̻���ǰ"
			elseif oip.flist(i).auction_cate_code = "99140200" then
				response.write "ĳ���Ϳ�ǰ"	
			elseif oip.flist(i).auction_cate_code = "99140300" then
				response.write "�ֹ�����/���㼱��"
			elseif oip.flist(i).auction_cate_code = "99140400" then
				response.write "����ٹ�/�ڽ�/Ȧ��"
			elseif oip.flist(i).auction_cate_code = "99140500" then
				response.write "Ű��Ʈ��ǰ"
			elseif oip.flist(i).auction_cate_code = "99140600" then
				response.write "�����μ�ǰ"		
			elseif oip.flist(i).auction_cate_code = "99140700" then
				response.write "���̵���ǰ"
			elseif oip.flist(i).auction_cate_code = "99140800" then
				response.write "��ļ�ǰ"	
			elseif oip.flist(i).auction_cate_code = "99140900" then
				response.write "��Ÿ��ǰ"										
			end if %><br>(<%= oip.flist(i).auction_cate_code %>)
			</td>
			<td align="center">
		<!--�������н��� -->	
				<input type="button" value="����" onclick="DelMe(frmBuyPrc<%=i%>,'<%= oip.flist(i).idx %>');">
		<!--�������г� -->
			</td>
    	</tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="7" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
        	<% if oip.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oip.StartScrollPage-1 %>')">[pre]</a>
	   		<% else %>
	    		[pre]
	   		<% end if %>
	
	    	<% for i=0 + oip.StartScrollPage to oip.FScrollCount + oip.StartScrollPage - 1 %>
	    		<% if i>oip.FTotalpage then Exit for %>
		    		<% if CStr(page)=CStr(i) then %>
		    		<font color="red">[<%= i %>]</font>
		    		<% else %>
		    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		    		<% end if %>
	    	<% next %>
	
	    	<% if oip.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
    		<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

<iframe frameboarder=0 height=0 width=0 name="view" id="view"></iframe>