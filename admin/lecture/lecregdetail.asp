<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/lecturecls.asp"-->
<%
dim idx,mode
dim olec

idx = request("idx")
mode = request("mode")

if idx="" then idx=0
set olec = new CLectureDetail
olec.GetLectureDetail idx

dim itemid,odetail
itemid = olec.Flinkitemid
set odetail = new CLecture
odetail.FRectItemID = itemid
odetail.GetLectureRegList

dim i
dim totno

totno =0
%>
<script language='javascript'>

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

function SendMsg(){
	if (!CheckSelected()){
		alert('�ϳ��̻� ���� �ϼž� �մϴ�.');
		return;
	}

	var ret = confirm('�޽����� �����ðڽ��ϱ�?');
	if (ret){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					refreshFrm.orderserial.value = refreshFrm.orderserial.value + frm.orderserial.value + ",";
				}
			}
		}
		var popwin = window.open('','refreshFrm','width=300 height=300');
		popwin.focus();
		refreshFrm.idx.value="<%=idx %>";
		refreshFrm.target = "refreshFrm";
		refreshFrm.action = "/admin/lecture/lecture_inputmsg.asp";
		refreshFrm.submit();
	}
}
</script>


<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td >���¸�</td>
	<td bgcolor="#FFFFFF"><% = olec.Flectitle %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�����</td>
	<td bgcolor="#FFFFFF"><% = olec.Flecturer %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >���º�</td>
	<td bgcolor="#FFFFFF">
		<% = olec.Flecsum %>
		<% if olec.Fmatinclude = "Y" then %>
		(��������)
		<% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����</td>
	<td bgcolor="#FFFFFF"><% = olec.Fmatsum %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >���ǱⰣ<br>(�ֱ�)</td>
	<td bgcolor="#FFFFFF"><% = olec.Flecperiod %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >���ǽð�</td>
	<td bgcolor="#FFFFFF"><% = olec.Flectime %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>�����Ͻ�</td>
	<td bgcolor="#FFFFFF">
		<% if Left(olec.Flecdate01,10)<>"1900-01-01" then %>
		1�� : <% = olec.Flecdate01 %>~<% = olec.Flecdate01_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate02,10)<>"1900-01-01" then %>
		2�� : <% = olec.Flecdate02 %>~<% = olec.Flecdate02_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate03,10)<>"1900-01-01" then %>
		3�� : <% = olec.Flecdate03 %>~<% = olec.Flecdate03_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate04,10)<>"1900-01-01" then %>
		4�� : <% = olec.Flecdate04 %>~<% = olec.Flecdate04_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate05,10)<>"1900-01-01" then %>
		5�� : <% = olec.Flecdate05 %>~<% = olec.Flecdate05_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate06,10)<>"1900-01-01" then %>
		6�� : <% = olec.Flecdate06 %>~<% = olec.Flecdate06_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate07,10)<>"1900-01-01" then %>
		7�� : <% = olec.Flecdate07 %>~<% = olec.Flecdate07_end %><br>
		<% end if %>
		<% if Left(olec.Flecdate08,10)<>"1900-01-01" then %>
		8�� : <% = olec.Flecdate08 %>~<% = olec.Flecdate08_end %><br>
		<% end if %>
	</td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<tr>
		<td align=right class="a" bgcolor="#FFFFFF">
		<input type="button" value="�޽��� ������" onClick="SendMsg();" class="button">
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#3d3d3d" class="a">
<tr bgcolor="#DDDDFF">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>�ֹ���ȣ</td>
	<td>����</td>
	<td>����</td>
	<td>���̵�</td>
	<td>����</td>
	<td>��ȭ</td>
	<td>�ڵ���</td>
	<td>�̸���</td>
	<td>�ֹ���</td>
	<td>������</td>
</tr>
<% for i=0 to odetail.FResultCount -1 %>
<%
if Not odetail.FItemList(i).IsCancel then
totno = totno + odetail.FItemList(i).Fitemno
end if
%>
<form name="frmBuyPrc_<%=i%>" method="post" action="" >
<input type="hidden" name=orderserial value=<%= odetail.FItemList(i).FOrderserial %>>
<tr bgcolor="#FFFFFF">
	<td>
		<% if odetail.FItemList(i).FIpkumdiv >=3 and Not(odetail.FItemList(i).IsCancel) then %>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		<% else %>
		<input type="checkbox" name="cksel" disabled>
		<% end if %>
	</td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FOrderserial %></font></td>
	<td><Font color=<%= odetail.FItemList(i).GetStateColor %> ><%= odetail.FItemList(i).GetStateName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyName %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserID %></font></td>
	<td align="center"><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).Fitemno %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyPhone %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FBuyHp %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FUserEmail %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FRegdate %></font></td>
	<td><Font color=<%= odetail.FItemList(i).IsCancelColor %> ><%= odetail.FItemList(i).FIpkumDate %></font></td>
</tr>
</form>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan=5></td>
	<td align="center"><%= totno %></td>
	<td colspan=6></td>
</tr>
</table>
<form name=refreshFrm method=post>
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="idx" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->