<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �������� ���� �𺰱�������
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,idx , i , shopid,zonename,racktype,unit,orderno,regdate ,isusing , zonegroup ,menupos
dim omanager ,managershopyn
	idx = requestCheckVar(request("idx"),10)
	menupos = requestCheckVar(request("menupos"),10)

set ozone = new czone_list
	ozone.frectidx = idx

set omanager = new czone_list
	omanager.frectzoneidx = idx
	
	'//�����ÿ��� ����
	if idx <> "" then		
		ozone.fzone_oneitem()
		
		if ozone.ftotalcount >0 then			
			shopid = ozone.FOneItem.fshopid		
			zonename = ozone.FOneItem.fzonename
			unit = ozone.FOneItem.funit			
			regdate = ozone.FOneItem.fregdate
			isusing = ozone.FOneItem.fisusing
			managershopyn = ozone.FOneItem.fmanagershopyn
			
			if managershopyn = "Y" then
				omanager.Getshopzonemanager()
			end if
		end if
	end if
	
%>

<script type="text/javascript">

	window.resizeTo(800, 500);
	
	function reg(){
		if (frm.shopid.value=='') {
			alert('������ ������ �ּ���');
			frm.zonename.focus();
			return;
		}
		
		if (frm.zonename.value=='') {
			alert('���׸��� �Է��� �ּ���');
			frm.zonename.focus();
			return;
		}
		
		if (frm.unit.value=='') {
			alert('���� ũ�⸦ �Է��� �ּ���');
			frm.unit.focus();			
			return;
		}
		
		if(frm.unit.value!=''){
			if (!IsDouble(frm.unit.value)){
				alert('���� ũ��� ���ڸ� �����մϴ�.');
				frm.unit.focus();
				return;
			}
		}	

		if (frm.isusing.value=='') {
			alert('��뿩�θ� ������ �ּ���');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='/admin/offshop/zone2/zone_process.asp';
		frm.mode.value = "zonereg";
		frm.submit();
	}

    // ���� ���� �˾�
	function popmanagerSelect(){
		var popmanagerSelect = window.open("/admin/offshop/zone2/pop_managerSelect.asp", "popmanagerSelect","width=600,height=400,scrollbars=yes,resizable=yes");
		popmanagerSelect.focus();
	}

	//�˾����� �Ŵ��� ���� �߰�
	function addSelectedmanager(empno,username){
		var lenRow = tablemanager.rows.length;

		// ������ ���� �ߺ��� ���� �˻�
		if(lenRow>1)	{
			for(l=0;l<document.all.empno.length;l++)	{
				if(document.all.empno[l].value==empno) {
					alert("�̹� ������ ����� �Դϴ�");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.empno.value==empno) {
					alert("�̹� ������ ����� �Դϴ�");
					return;
				}
			}
		}

		// ���߰�
		var oRow = tablemanager.insertRow(lenRow);
		oRow.onmouseover=function(){tablemanager.clickedRowIndex=this.rowIndex};

		// ���߰� (�̸�,������ư)
		var oCell1 = oRow.insertCell(0);		
		var oCell3 = oRow.insertCell(1);

		oCell1.innerHTML = username + "<input type='hidden' name='empno' value='" + empno + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdmanager()' align=absmiddle>";
	}

	// ���û���
	function delSelectdmanager(){
	    
		if(confirm("������ ����ڸ� �����Ͻðڽ��ϱ�?"))
			tablemanager.deleteRow(tablemanager.clickedRowIndex);
	}
			
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		�� <font color="red">[�߿�] </font>���峻 ������ ����ǰų� ��������,
		<br>���� ������ �������� ���ð�, ������ ��������, ���� ����ϼ���.
		<br>���� ������ ���� ����� �������� ������ ��� �Ͻǰ��,
		<br>���� �������� ��ϵǾ��� ��ǰ���� ��� ���� �������� ����Ǵ� ������ �߻��˴ϴ�
	</td>
	<td align="right">
		<input type="button" value="�űԵ��" class="button" onclick="location.href='?menupos=<%=request("menupos")%>';">&nbsp;&nbsp;
		<input type="button" value="â�ݱ�" class="button" onclick="window.close();">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method="post" style="margin:0px;" >
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
	<td align="center">��ȣ<br></td>
	<td>
		<%=idx%><input type="hidden" name="idx" value="<%=idx%>">
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">SHOP</td>
	<td>
		<% drawSelectBoxOffShop "shopid",shopid %>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td align="center">���׸�</td>
	<td>
		<input type="text" name="zonename" value="<%=zonename%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">����ũ��</td>
	<td>
		<input type="text" name="unit" value="<%=unit%>" size=5 maxlength=5> ex)1
		<p>�� �ش������� ����� ����Ͻðų�, ���������� ���ϽŴ�� �����ؼ� ����Ͻø� �˴ϴ�
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">���峻�����<br></td>
	<td>
		<table border="0" cellspacing="0" class="a">
		<tr>
			<td>
			    <table name='tablemanager' id='tablemanager' class=a>
			    <% if managershopyn = "Y" then %>
			        <% for i=0 to omanager.FResultCount-1 %>
			        <tr onMouseOver='tablemanager.clickedRowIndex=this.rowIndex'>
				    	<td>
				    	    <%= omanager.FItemList(i).fusername %>
				    	    <input type='hidden' name='empno' value='<%= omanager.FItemList(i).fempno %>'></td>  
				    	<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdmanager()' align=absmiddle></td>   
			        </tr>
			        <% next %>
			    <% end if %>
			    </table>
			</td>
			<td valign="bottom"><input type="button" class='button' value="�߰�" onClick="popmanagerSelect()"></td>
		<tr>
	    </table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">��뿩��<br></td>
	<td>
		<select name="isusing">
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="����" class="button" onclick="reg();">
	</td>
</tr>
</table>	
</form>

<%
set ozone = nothing
set omanager = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
