<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->

<%
dim idx
dim ofranchulgojungsan, shopid

idx = RequestCheckvar(request("idx"),10)

if idx="" then idx="0"

set ofranchulgojungsan = new CFranjungsan
ofranchulgojungsan.FRectidx = idx
ofranchulgojungsan.getOneFranJungsan

Dim defaultYYYY, defaultMM, defaultShopDiv

IF idx="0" THen
    defaultYYYY = Left(DateAdd("m",-1,now()),4)
    defaultMM   = Mid(DateAdd("m",-1,now()),6,2)
    defaultShopDiv = ""
ELSE
    defaultYYYY = ""
    defaultMM   = ""
    defaultShopDiv = ""
END IF
%>
<script language='javascript'>
function SaveInfo(frm){
	if (frm.title.value.length<1){
		alert('Title�� �Է��ϼ���');
		frm.title.focus();
		return;
	}

	if (frm.shopdiv.value.length<1){
		alert('������ �Է��ϼ���');
		frm.shopdiv.focus();
		return;
	}

	if (frm.diffKey.value.length<1){
	    alert('���� ������ �Է��ϼ���');
		frm.diffKey.focus();
		return;
	}

<% if idx="0" then %>
	if (frm.shopid.value.length<1){
		alert('����ID�� �Է��ϼ���');
		frm.shopid.focus();
		return;
	}

	if (frm.totalbuycash.value.length<1){
		alert('�� ���԰��� �Է��ϼ���');
		frm.totalbuycash.focus();
		return;
	}

	if (frm.totalsuplycash.value.length<1){
		alert('�� ���ް��� �Է��ϼ���');
		frm.totalsuplycash.focus();
		return;
	}
<% end if %>

/*
	if (frm.totalsum.value.length<1){
		alert('�� ����ݾ��� �Է��ϼ���');
		frm.totalsum.focus();
		return;
	}


	if ((!frm.statecd[0].checked)&&(!frm.statecd[1].checked)&&(!frm.statecd[2].checked)&&(!frm.statecd[3].checked)){
		alert('���¸� �����ϼ���.');
		frm.statecd[0].focus();
		return;
	}
*/

	if (confirm('���� �Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function PopSegumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		//if (confirm('������ : ' + comp.value + ' OK?')){
		//	frm.idx.value = iidx;
		//	frm.mode.value = "segumil";
		//	frm.submit();
		//}
	};
}

function PopIpgumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		//if (confirm('�Ա��� : ' + comp.value + ' OK?')){
		//	frm.idx.value = iidx;
		//	frm.mode.value="ipkumil";
		//	frm.submit();
		//}
	};
}


function changeState(state)
{
	var f = document.frm;

	switch (state)
	{
	case "0":
		var msg = "���������� �����Ͻðڽ��ϱ�?";
		break;
	case "1":
		var msg = "��üȮ�������� �����Ͻðڽ��ϱ�?";
		break;
	case "3":
		var msg = "��üȮ�οϷ�� �����Ͻðڽ��ϱ�?";
		break;
	case "7":
		var msg = "�ԱݿϷ�� �����Ͻðڽ��ϱ�?";
		if (f.ipkumdate.value.length!=10)
		{
			alert("�Ա����� �Է��Ͻʽÿ�.");
			return;
		}
		break;
	}

	if (confirm(msg))
	{
		f.mode.value = "changeState";
		f.stateCd.value = state;
		f.submit();
	}
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=post action="meaipchulgojungsan_process.asp">
	<input type=hidden name="idx" value="<%= ofranchulgojungsan.FOneItem.Fidx %>">
	<% if idx="0" then %>
	<input type=hidden name="mode" value="addmaster">
	<% else %>
	<input type=hidden name="mode" value="modimaster">

	<input type="hidden" name="stateCd" value="">
	<% end if %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>IDX</td>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fidx %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">������ID</td>
		<% if idx="0" then %>
		<td bgcolor="#FFFFFF" ><% drawSelectBoxOffShopNot000 "shopid",shopid %></td>
		<% else %>
		<td bgcolor="#FFFFFF" ><%= ofranchulgojungsan.FOneItem.Fshopid %></td>
		<% end if %>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<% if idx="0" then %>
		<td bgcolor="#FFFFFF" ><% call DrawYMBox(defaultYYYY,defaultMM) %></td>
		<% else %>
		<td bgcolor="#FFFFFF" ><% call DrawYMBox(Left(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),4),Right(NULL2Blank(ofranchulgojungsan.FOneItem.FYYYYMM),2)) %></td>
		<% end if %>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		<td bgcolor="#FFFFFF" >
			<% if idx="0" then %>
			<% Call DrawShopDivBox(defaultShopDiv) %>
			/
			<select class="select" name="divcode">
				<option value="GC">���ͺ�
				<option value="ET">��Ÿ���
			</select>


			<% else %>
			<% Call DrawShopDivBox(ofranchulgojungsan.FOneItem.FShopDiv) %>
			/
			<font color="<%= ofranchulgojungsan.FOneItem.GetDivCodeColor %>"><%= ofranchulgojungsan.FOneItem.GetDivCodeName %></font>


			<% end if %>


		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	    <td bgcolor="#FFFFFF" >
	    <% if idx="0" then %>
	    <input type="text" name="diffKey" maxlength="2" class="text">
	    <% else %>
	    <input type="text" name="diffKey" value="<%= ofranchulgojungsan.FOneItem.FdiffKey %>" size="2" maxlength="2" class="text">
	    <% end if %>
	    </td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">Title</td>
		<td bgcolor="#FFFFFF" >
			<input type="text" class="text" name=title value="<%= ofranchulgojungsan.FOneItem.Ftitle %>" size="30" maxlength="30" <%If ofranchulgojungsan.FOneItem.Fstatecd>="4" Then %>readOnly<%End If %> >
			(ex) OO�� 4�� 1�� ��ǰ��
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѼҺ��ڰ�</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>

			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsellcash,0) %>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ѹ��԰�</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
			<input type=text name=totalbuycash value="" size=9 maxlength=9 style="border:1px #999999 solid; text-align=right">
			<font color="#AAAAAA">(�ҿ� ���:����)</font>
			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalbuycash,0) %>
			<font color="#AAAAAA">(��ü�κ��� ���޹��� ��ǰ����)</font>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ѱ��ް�</td>
		<td bgcolor="#FFFFFF">
			<% if idx="0" then %>
			<input type=text name=totalsuplycash value="" size=9 maxlength=9 style="border:1px #999999 solid; text-align=right">
			<font color="#AAAAAA">(������ ������ ��ǰ����)</font>
			<% else %>
			<%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsuplycash,0) %>
			<font color="#AAAAAA">(������ ������ ��ǰ����)</font>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�ѹ���ݾ�</td>
		<td bgcolor="#FFFFFF">
		    <%= FormatNumber(ofranchulgojungsan.FOneItem.Ftotalsum,0) %> &nbsp;(��꼭 ���� �ݾ�)
		    <!--
			<input type="text" class="text" name="totalsum" value="<%= ofranchulgojungsan.FOneItem.Ftotalsum %>" size="10" maxlength="9" style="text-align=right"> (��꼭 ���� �ݾ�)
			-->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��꼭������</td>
		<td bgcolor="#FFFFFF">
		    <%= ofranchulgojungsan.FOneItem.Ftaxdate %>
		    <!--
			<input type="text" class="text" name="taxdate" value="<%= ofranchulgojungsan.FOneItem.Ftaxdate %>" size="10" maxlength="10">
			<a href="javascript:PopSegumil(frm,'<%= ofranchulgojungsan.FOneItem.Fidx %>',frm.taxdate);"><img src="/images/calicon.gif" align="absmiddle" border="0"></a>
			-->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�Ա���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="ipkumdate" value="<%= ofranchulgojungsan.FOneItem.Fipkumdate %>" size="10" maxlength="10" readonly>
		<%if (ofranchulgojungsan.FOneItem.Fstatecd="4") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
			<a href="javascript:PopIpgumil(frm,'<%= ofranchulgojungsan.FOneItem.Fidx %>',frm.ipkumdate);"><img src="/images/calicon.gif" align="absmiddle" border="0"></a>
		<%end if %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		<td bgcolor="#FFFFFF" >
		<font color="<%= ofranchulgojungsan.FOneItem.GetStateColor %>"><%= ofranchulgojungsan.FOneItem.GetStateName %></font>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="0") then %>
		==&gt; <input type="button" class="button" onclick="changeState('1');" value="��üȮ�������� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="1") then %>
		==&gt; <input type="button" class="button" onclick="changeState('3');" value="��üȮ�οϷ�� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		==&gt; <input type="button" class="button" onclick="changeState('7');" value="�ԱݿϷ�� ����">
		<% else %>
		<% end if %>

		<% if (ofranchulgojungsan.FOneItem.Fstatecd="1") or (ofranchulgojungsan.FOneItem.Fstatecd="3") then %>
		<input type="button" class="button" onclick="changeState('0');" value="���������� ����">
		<% elseif (ofranchulgojungsan.FOneItem.Fstatecd="4") then %>
		<input type="button" class="button" onclick="changeState('0');" value="���������� ����">
		<% else %>

	    <% end if %>
		<!--
			<input type=radio name=statecd value="0" <% if ofranchulgojungsan.FOneItem.Fstatecd="0" then response.write "checked" %>>������
			<input type=radio name=statecd value="1" <% if ofranchulgojungsan.FOneItem.Fstatecd="1" then response.write "checked" %>>��üȮ����
			<input type=radio name=statecd value="4" <% if ofranchulgojungsan.FOneItem.Fstatecd="4" then response.write "checked" %>>��꼭����
			<input type=radio name=statecd value="7" <% if ofranchulgojungsan.FOneItem.Fstatecd="7" then response.write "checked" %>>�ԱݿϷ�
		-->
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">��Ÿ����</td>
		<td bgcolor="#FFFFFF" >
			<textarea name="etcstr" class="textarea" cols="86" rows="8"><%= ofranchulgojungsan.FOneItem.Fetcstr %></textarea>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">���ʵ����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregusername %>(<%= ofranchulgojungsan.FOneItem.Freguserid %>)</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">����ó����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Ffinishusername %>(<%= ofranchulgojungsan.FOneItem.Ffinishuserid %>)</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
		<td bgcolor="#FFFFFF"><%= ofranchulgojungsan.FOneItem.Fregdate %></td>
	</tr>
	<tr>
		<td colspan=2 align=center bgcolor="#FFFFFF">
		<%If idx="0" Then %>
			<input type="button" class="button" value="��������" onclick="SaveInfo(frm);">
		<% else %>
			<input type="button" class="button" value="����" onclick="SaveInfo(frm);">
		<%End If %>

		</td>
	</tr>
	</form>
</table>
* ������ ����/�Ա���/��Ÿ�޸� ����˴ϴ�.
<%
set ofranchulgojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->