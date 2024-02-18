<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2015.11.05 �ѿ�� ����
'	Description : ���� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/pack_cls.asp"-->

<%
Dim orderserial, i, midx, userid, tmpmidx
	orderserial = requestcheckvar(request("orderserial"),11)
	midx = getNumeric(requestcheckvar(request("midx"),10))

dim vtitle, vmessage

if orderserial="" then
	response.write "<script type='text/javascript'>alert('�ֹ���ȣ�� �����ϴ�.'); self.close();</script>"
	dbget.close()	:	response.end
end if

dim cpacksum
set cpacksum = new Cpack
	cpacksum.FRectOrderSerial = orderserial
	cpacksum.Getpojang_itemlist()

%>

<script type="text/javascript">

function detailview(orderserial, midx){
	location.replace("/cscenter/pack/pojang_view.asp?orderserial="+orderserial+"&midx="+midx);
}

function editproc(midx){
	if (midx==''){
		alert('�ϷĹ�ȣ�� �����ϴ�.');
		return;
	}

	if (pojangfrm.title.value == '' || GetByteLength(pojangfrm.title.value) > 60){
		alert("���� ������� ���ų� ���ѱ��̸� �ʰ��Ͽ����ϴ�. 60�� ���� �ۼ� �����մϴ�.");
		pojangfrm.title.focus();
		return;
	}
	if (pojangfrm.message.value != '' && GetByteLength(pojangfrm.title.value) > 100){
		alert("���� �޼����� ���ѱ��̸� �ʰ��Ͽ����ϴ�. 100�� ���� �ۼ� �����մϴ�.");
		pojangfrm.message.focus();
		return;
	}

	pojangfrm.mode.value='editpojang';
	pojangfrm.midx.value=midx;
	pojangfrm.action = "/cscenter/pack/pojang_process.asp";
	pojangfrm.submit();
	return;
}

</script>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		�ֹ���ȣ : <%= orderserial %>
	</td>
	<td align="right">
	</td>
</tr>
<tr>
	<td align="left">
	</td>
</tr>
</tr>
</table>

<br>
<font color="red">�ؼ������峻��</font>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�ѹڽ��� : <b><%= cpacksum.Fpackcnt %></b> / �ѻ�ǰ�����հ� : <b><%= cpacksum.Fpackitemcnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����</td>
	<td>����޼���</td>
	<td>��ǰ��<br>�հ�</td>
	<td>������</td>
	<td>��������</td>
	<td>�ڽ���ȣ<br>(������ X)</td>
	<td>���</td>
</tr>
<% if cpacksum.FresultCount>0 then %>
	<% for i=0 to cpacksum.FresultCount-1 %>
		<% if cstr(midx)=cstr(cpacksum.FItemList(i).fmidx) then %>
			<%
			vtitle = cpacksum.FItemList(i).ftitle
			vmessage = cpacksum.FItemList(i).fmessage
			%>
			<tr align="center" bgcolor="orange" >
		<% else %>
			<tr align="center" bgcolor="#FFFFFF" >
		<% end if %>

		<% if tmpmidx="" or cstr(tmpmidx)<>cstr(cpacksum.FItemList(i).fmidx)  then %>
			<td>
				<%= chrbyte(cpacksum.FItemList(i).ftitle,10,"Y") %>
			</td>
			<td>
				<%= chrbyte(cpacksum.FItemList(i).fmessage,10,"Y") %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fpackitemcnt %>
			</td>
			<td>
				<%= FormatDate(cpacksum.FItemList(i).fregdate,"0000.00.00") %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fcancelyn %>
			</td>
			<td>
				<%= cpacksum.FItemList(i).fmidx %>
			</td>
			<td>
				<input type="button" onclick="detailview('<%= orderserial %>','<%= cpacksum.FItemList(i).fmidx %>');" value="����" class="button">
			</td>
		<% end if %>

		<% tmpmidx = cpacksum.FItemList(i).fmidx %>

		<% if cstr(tmpmidx)=cstr(cpacksum.FItemList(i).fmidx) then %>
			</tr>
			<tr>
				<td colspan=7 align="right" bgcolor="#FFFFFF">
					<table width="900" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr bgcolor="#FFFFFF" >
						<td width=55>
							<img src="<%= cpacksum.FItemList(i).FImageSmall %>" width=50 height=50>
						</td>
						<td width=65>
							��ǰ�ڵ�:
							<br><%= cpacksum.FItemList(i).FItemID %>
						</td>
						<td width=130>
							�귣���:
							<br><%= cpacksum.FItemList(i).FBrandName %>
						</td>
						<td>
							��ǰ��:
							<br><%= cpacksum.FItemList(i).FItemName %>
						</td>
						<td width=170>
							<% if cpacksum.FItemList(i).FItemOptionName<>"" then %>
								�ɼǸ�:
								<br><%= cpacksum.FItemList(i).FItemOptionName %>
							<% end if %>
						</td>
						<td width=70>
							����: <%= cpacksum.FItemList(i).FItemEa %>
						</td>
						<td width=50>
							����: <%= cpacksum.FItemList(i).fcancelyn %>
						</td>
					</tr>
					</table>
				</td>
		<% end if %>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<% if midx<>"" then %>
	<br>
	<font color="red">�ؼ����������</font>
	<br>
	<form name="pojangfrm" method="post" action="" style="margin:0px;">
	<input type="hidden" name="mode">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<tr>
		<td bgcolor="#e1e1e1" align="center">�ڽ���ȣ</td>
		<td bgcolor="#FFFFFF">
			<%= midx %>
			<input type="hidden" name="midx">
		</td>
	</tr>
	<tr>
		<td bgcolor="#e1e1e1" align="center">�����</td>
		<td bgcolor="#FFFFFF">
			<input type="text" name="title" value="<%= vtitle %>" size="100">
		</td>
	</tr>
	<tr>
		<td bgcolor="#e1e1e1" align="center">����޼���</td>
		<td bgcolor="#FFFFFF">
			<textarea name="message" style="width:600px;" rows="5"><%= vmessage %></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" align="center" colspan=2>
			<input type="button" onclick="editproc('<%= midx %>');" value="����" class="button">
		</td>
	</tr>
	</table>
	</form>
<% end if %>

<%
set cpacksum = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
