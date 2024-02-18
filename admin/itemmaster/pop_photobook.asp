<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/item_photobook_cls.asp"-->
<%
dim oitem, oitem1, page, i, vItemid
page = request("page")
vItemid = request("itemid")

If page = "" Then
	page = 1
End If

set oitem = new CItemPhoto
oitem.FPageSize         = 10
oitem.FCurrPage         = page
oitem.GetPhotoItemList()


Dim vSelect1, vSelect2
vSelect1 = "<select name='pcode' class='select'>" & _
			"<option value=''>- ���� -</option>" & _
			"<option value='550000094'>����� 5x5</option>" & _
			"<option value='550000001'>����� 6x6</option>" & _
			"<option value='550000002'>����� 8x8</option>" & _
			"<option value='550000095'>����� 10x10</option>" & _
			"<option value='550000186'>ĳ���� ����� 6x6</option>" & _
			"<option value='550000187'>ĳ���� ����� 8x8</option>" & _
			"<option value='550000034'>Ź��� Ķ���� 6x8</option>" & _
			"<option value='550000195'>ĳ���� Ķ���� 6x10</option>" & _
			"</select>"
			
vSelect2 = "<select name='tplname' class='select'>" & _
			"<option value=''>- ���� -</option>" & _
			"<option value='photobook'>photobook</option>" & _
			"<option value='calendar'>calendar</option>" & _
			"</select>"
%>

<script type="text/javascript"> 
function addtr(obj) { 

	oRow = document.createElement('tr');
	oCel0 = document.createElement('td');
	oCel1 = document.createElement('td');
	oCel2 = document.createElement('td');
	oCel3 = document.createElement('td');
	oCel4 = document.createElement('td');
	
	oCel0['bgColor']='#FFFFFF';
	oCel1['bgColor']='#FFFFFF';
	oCel2['bgColor']='#FFFFFF';
	oCel3['bgColor']='#FFFFFF';
	oCel4['bgColor']='#FFFFFF'; 
	
	oCel0.innerHTML="<input type='text' name='itemid' value='' size='7'>";
	oCel1.innerHTML="<input type='text' name='itemoption' value='' size='10'>";
	oCel2.innerHTML="<input type='text' name='tplcode' value='' size='7'>";
	oCel3.innerHTML="<%=vSelect1%>";
	oCel4.innerHTML="<%=vSelect2%>";

	oRow.appendChild(oCel0);
	oRow.appendChild(oCel1);
	oRow.appendChild(oCel2);
	oRow.appendChild(oCel3);
	oRow.appendChild(oCel4);

	document.getElementById('FAM_TABLE').children(0).appendChild(oRow);
} 

function goSumit()
{
	var f = document.frm;
	var totalcnt = document.getElementsByName("itemid").length;

	if(totalcnt == 1)
	{
		if(f.itemid.value == "")
		{
			alert("��ǰ�ڵ带 �Է��ϼ���.");
			return;
		}
		else
		{
			if(isNaN(f.itemid.value))
			{
				alert("��ǰ�ڵ带 ���ڷθ� �Է��ϼ���.");
				f.itemid.value = "";
				f.itemid.focus();
				return;
			}
			
			if(f.itemoption.value == "" || f.tplcode.value == "" || f.pcode.value == "" || f.tplname.value == "")
			{
				alert("��ǰ�ڵ带 �Է��� �κп��� ��� ���� �Է��ؾ��մϴ�.");
				return;
			}
		}
		
		if(confirm("��ǰ�ڵ�� �ɼ��ڵ尡 �ٸ��� �ԷµǾ����ϱ�?\n�߸��Է½� �����Ұ��Դϴ�.") == true) {
			f.submit();
		}
		else
		{
			return;
		}
	}
	else
	{
		for(var i=0; i<totalcnt; i++)
		{
			if(!(f.itemid[i].value == "" && f.itemoption[i].value == "" && f.tplcode[i].value == "" && f.pcode[i].value == "" && f.tplname[i].value == ""))
			{
				if(f.itemid[i].value == "")
				{
					alert("��ǰ�ڵ带 �Է��ϼ���.");
					return;
				}
				else
				{
					if(isNaN(f.itemid[i].value))
					{
						alert("��ǰ�ڵ带 ���ڷθ� �Է��ϼ���.");
						f.itemid[i].value = "";
						f.itemid[i].focus();
						return;
					}
					
					if(f.itemoption[i].value == "" || f.tplcode[i].value == "" || f.pcode[i].value == "" || f.tplname[i].value == "")
					{
						alert("��ǰ�ڵ带 �Է��� �κп��� ��� ���� �Է��ؾ��մϴ�.");
						return;
					}
				}
			}
		}
		if(confirm("��ǰ�ڵ�� �ɼ��ڵ尡 �ٸ��� �ԷµǾ����ϱ�?\n�߸��Է½� �����Ұ��Դϴ�.") == true) {
			f.submit();
		}
		else
		{
			return;
		}
	}
}
</script> 

<table cellpadding="0" cellspacing="0" class="a">
<tr>
	<td><b><font size="2">����� ���ø��ڵ� �Է�â</font></b></td>
</tr>
<tr>
	<td style="padding:5 0 5 0;">
		�� ��ǰ�ڵ�� ��ǰ�ɼ��ڵ�� �����Ұ��ϴ� �����Ͽ� �Է��ϼ���.
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type='button' name='add' value='�����Է�' class="button" onclick='location.href="?page=<%=page%>";'>
	</td>
</tr>
</table>

<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id='FAM_TABLE'>
<form name="frm" methopd="post" action="pop_photobook_proc.asp">
<tr>
	<td align="center" width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
	<td align="center" width="80" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ɼ��ڵ�</td>
	<td align="center" width="60" bgcolor="<%= adminColor("tabletop") %>">�����ڵ�</td>
	<td align="center" width="150" bgcolor="<%= adminColor("tabletop") %>">���ø��ڵ��ڵ�</td>
	<td align="center" width="150" bgcolor="<%= adminColor("tabletop") %>">���ø���</td>
</tr>
<%
If vItemid <> "" Then
	
	Response.Write "<input type=""hidden"" name=""gubun"" value=""update"">"
	
	set oitem1 = new CItemPhoto
	oitem1.FRectItemId = vItemid
	oitem1.GetPhotoTempleteList()
	
	for i=0 to oitem1.FresultCount
%>
	<tr>
		<td bgcolor="#FFFFFF"><input type="text" name="itemid" value="<%= oitem1.FItemList(i).Fitemid %>" size="7" readonly></td>
		<td bgcolor="#FFFFFF"><input type="text" name="itemoption" value="<%= oitem1.FItemList(i).Fitemoption %>" size="10" readonly></td>
		<td bgcolor="#FFFFFF"><input type="text" name="tplcode" value="<%= oitem1.FItemList(i).Ftplcode %>" size="7"></td>
		<td bgcolor="#FFFFFF">
			<select name="pcode" class="select">
				<option value="">- ���� -</option>
				<option value="550000094" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000094","selected","")%>>����� 5x5</option>
				<option value="550000001" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000001","selected","")%>>����� 6x6</option>
				<option value="550000002" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000002","selected","")%>>����� 8x8</option>
				<option value="550000095" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000095","selected","")%>>����� 10x10</option>
				<option value="550000186" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000186","selected","")%>>ĳ���� ����� 6x6</option>
				<option value="550000187" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000187","selected","")%>>ĳ���� ����� 8x8</option>
				<option value="550000034" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000034","selected","")%>>Ź��� Ķ���� 6x8</option>
				<option value="550000195" <%=CHKIIF(oitem1.FItemList(i).Fpcode="550000195","selected","")%>>ĳ���� Ķ���� 6x10</option>
			</select>
		</td>
		<td bgcolor="#FFFFFF">
			<select name="tplname" class="select">
				<option value="">- ���� -</option>
				<option value="photobook" <%=CHKIIF(oitem1.FItemList(i).Ftplname="photobook","selected","")%>>photobook</option>
				<option value="calendar" <%=CHKIIF(oitem1.FItemList(i).Ftplname="calendar","selected","")%>>calendar</option>
			</select>
		</td>
	</tr>
<%
	next
	
	set oitem1 = nothing
Else
	Response.Write "<input type=""hidden"" name=""gubun"" value=""insert"">"
%>
<tr>
	<td bgcolor="#FFFFFF"><input type="text" name="itemid" value="" size="7"></td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemoption" value="" size="10"></td>
	<td bgcolor="#FFFFFF"><input type="text" name="tplcode" value="" size="7"></td>
	<td bgcolor="#FFFFFF">
		<select name="pcode" class="select">
			<option value="">- ���� -</option>
			<option value="550000094">����� 5x5</option>
			<option value="550000001">����� 6x6</option>
			<option value="550000002">����� 8x8</option>
			<option value="550000095">����� 10x10</option>
			<option value="550000186">ĳ���� ����� 6x6</option>
			<option value="550000187">ĳ���� ����� 8x8</option>
			<option value="550000034">Ź��� Ķ���� 6x8</option>
			<option value="550000195">ĳ���� Ķ���� 6x10</option>
		</select>
	</td>
	<td bgcolor="#FFFFFF">
		<select name="tplname" class="select">
			<option value="">- ���� -</option>
			<option value="photobook">photobook</option>
			<option value="calendar">calendar</option>
		</select>
	</td>
</tr>
<%
End If
%>
</form>
</table>
<br>
<table width="100%" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td><input type='button' name='add' value='���߰�' size='5' onclick='addtr(this)'></td>
	<td align="right"><input type="button" value="��  ��" onClick="goSumit()"></td>
</tr>
</table>
<br><br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		Total : <b><%= oitem.FTotalCount%></b>
		&nbsp;
		������ : <b><%= page %> /<%=  oitem.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td width="50">����</td>
</tr>
<% if oitem.FresultCount<1 then %>
    <tr bgcolor="#FFFFFF">
    	<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
<% if oitem.FresultCount > 0 then %>
<% for i=0 to oitem.FresultCount-1 %>
<tr class="a" height="25" bgcolor="#FFFFFF">
	<td align="center"><%= oitem.FItemList(i).Fitemid %></td>
	<td><%= oitem.FItemList(i).Fitemname %></td>
	<td align="center"><input type="button" value="����" class="button" onClick="location.href='?page=<%=page%>&itemid=<%= oitem.FItemList(i).Fitemid %>';"></td>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<% if oitem.HasPreScroll then %>
		<a href="?page=<%= oitem.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oitem.StartScrollPage to oitem.FScrollCount + oitem.StartScrollPage - 1 %>
			<% if i>oitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oitem.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<% set oitem = nothing %>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->