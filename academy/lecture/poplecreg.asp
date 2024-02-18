<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ�
' History : 2010.05.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_lecturecls.asp"-->

<%
dim lec_idx , i , lecOption
lec_idx = RequestCheckvar(request("lec_idx"),10)
lecOption = RequestCheckvar(request("lecOption"),4)

dim olecture
set olecture = new CLecture
	olecture.FRectIdx = lec_idx
	'olecture.FRectLecOpt = lecOption

	if lec_idx<>"" then
		olecture.GetOneLecture
	end if
%>
<script language='javascript'>

	// ���� �ڸ��� ǥ��
	function plusComma(num){
		if (num < 0) { num *= -1; var minus = true}
		else var minus = false

		var dotPos = (num+"").split(".")
		var dotU = dotPos[0]
		var dotD = dotPos[1]
		var commaFlag = dotU.length%3

		if(commaFlag) {
			var out = dotU.substring(0, commaFlag)
			if (dotU.length > 3) out += ","
		}
		else var out = ""

		for (var i=commaFlag; i < dotU.length; i+=3) {
			out += dotU.substring(i, i+3)
			if( i < dotU.length-3) out += ","
		}

		if(minus) out = "-" + out
		if(dotD) return out + "." + dotD
		else return out
	}
    
    function chgEntryDetail(comp){
        var cnt = comp.value;
		var ttlsumid = document.all["htmlttlsum"];
		
		var matinclude_yn = eval("frmlec.matinclude_yn");
		var mat_cost = eval("frmlec.mat_cost");
		var mat_buying_cost = eval("frmlec.mat_buying_cost");

		var lec_cost = eval("frmlec.lec_cost");
		var buying_cost = eval("frmlec.buying_cost");

		var itemno = eval("frmlec.itemea").value;

		var sellprice = eval("frmlec.sellprice").value;
		var ttlsumvalue = itemno*1*sellprice;
		var soldoutflagform = eval("frmlec.soldoutflag");
		var itemsubtotalsumfrm = eval("frmlec.itemsubtotalsum");

		itemsubtotalsumfrm.value = ttlsumvalue;
        
        if (matinclude_yn.value == "C") {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>����";
		} else {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>����";
		}
    }
    
	// ������ 6���ڿ� ���� ���̺� ǥ��
	function ShowEntryDetail(comp){

		var cnt = comp.value;
		var ttlsumid = document.all["htmlttlsum"];

		var matinclude_yn = eval("frmlec.matinclude_yn");
		var mat_cost = eval("frmlec.mat_cost");
		var mat_buying_cost = eval("frmlec.mat_buying_cost");

		var lec_cost = eval("frmlec.lec_cost");
		var buying_cost = eval("frmlec.buying_cost");

		var itemno = eval("frmlec.itemea").value;

		var sellprice = eval("frmlec.sellprice").value;
		var ttlsumvalue = itemno*1*sellprice;
		var soldoutflagform = eval("frmlec.soldoutflag");
		var itemsubtotalsumfrm = eval("frmlec.itemsubtotalsum");

		itemsubtotalsumfrm.value = ttlsumvalue;

		for (i=0;i<cnt;i++){
			document.all["entry"+(i)].style.display="";
		}

		for (i=3;i>=cnt;i--){
			document.all["entry"+(i)].style.display="none";
		}

		if (matinclude_yn.value == "C") {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>����";
		} else {
			ttlsumid.innerHTML = "<b>" + plusComma(ttlsumvalue) + "</b><br>����";
		}

		//RecalcuSubTotal();

		//��Ŀ���̵�
		//eval("baguniFrm.buy_name").focus();
	}


	// �� �˻� �� ����
	function SaveItem()
	{
		var frm = document.frmlec;

		if(!frm.lecOption.value)
		{
			alert("���½ð��� �������ֽʽÿ�.");
			frm.lecOption.focus();
			return;
		} else if(frm.lecOption.options[frm.lecOption.selectedIndex].id=="S") {
			alert("������ ���´� ������û�� �� �� �����ϴ�.");
			frm.lecOption.focus();
			return;
		}

		if(!frm.buy_name.value)
		{
			alert("�ֹ����� �̸��� �Է����ֽʽÿ�.");
			frm.buy_name.focus();
			return;
		}

		if(!(frm.buy_phone1.value&&frm.buy_phone2.value&&frm.buy_phone3.value))
		{
			alert("�ֹ����� ��ȭ��ȣ�� �Է����ֽʽÿ�.");
			frm.buy_phone1.focus();
			return;
		}

		if(!(frm.buy_hp1.value&&frm.buy_hp2.value&&frm.buy_hp3.value))
		{
			alert("�ֹ����� �޴�����ȣ�� �Է����ֽʽÿ�.");
			frm.buy_hp1.focus();
			return;
		}

		if(!frm.buy_email.value)
		{
			alert("�ֹ����� �̸����� �Է����ֽʽÿ�.");
			frm.buy_email.focus();
			return;
		}
        <% If Not olecture.FOneItem.isWeClass Then %>
		for(i=1;i<frm.itemea.value;i++)
		{
			if(!frm['entryname' + i].value)
			{
				alert("������#" + (i+1) + "�� �̸��� �Է����ֽʽÿ�.");
				frm['entryname' + i].focus();
				return;
			}

			if(!(frm['entry' + i + '_hp1'].value&&frm['entry' + i + '_hp2'].value&&frm['entry' + i + '_hp3'].value))
			{
				alert("������#" + (i+1) + "�� ����ó�� �Է����ֽʽÿ�.");
				frm['entry' + i + '_hp1'].focus();
				return;
			}
		}
        <% end if %>
        
        if ((frm.paymethod.value=="7")&&(frm.lecOption.value!="0000")){
            alert('��ü���� �ֹ� ������ ��Ͻ� ���½ð�=���(�⺻��)���� �����ϼ���.');
            return;
        }
        
        <% If olecture.FOneItem.isWeClass Then %>
        if (frm.wantstudyName.value.length<1){
    		alert('�ֹ��� ��ü(��ȣȸ)���� �Է��Ͻñ� �ٶ��ϴ�.');
    		frm.wantstudyName.focus();
    		return;
    	}
	
    	if (frm.wantstudyPlace.value.length<1){
    		alert('������� �Է��Ͻñ� �ٶ��ϴ�.');
    		frm.wantstudyPlace.focus();
    		return;
    	}
    	if (!(frm.wantstudyWho[0].checked) && !(frm.wantstudyWho[1].checked) && !(frm.wantstudyWho[2].checked) && !(frm.wantstudyWho[3].checked)){
    		alert('���Ǵ���� �����Ͻñ� �ٶ��ϴ�.');
    		return;
    	}
    	<% End If %>
	
		// ����
		if (confirm('���� ���� ��� �Ͻðڽ��ϱ�?')){
		    frm.submit();
		}
	}
	

	// ���̵� ã��
	function popSrcId()
	{
		window.open("popSearchId.asp", "popId", "width=418,height=300,scrollbars=yes")
	}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmlec" method="POST" action="DoPopLecReg.asp">
<input type="hidden" name="lec_idx" value="<%=lec_idx%>">
<input type="hidden" name="lec_title" value="<%=olecture.FOneItem.Flec_title%>">

<input type="hidden" name="lec_cost" value="<%=olecture.FOneItem.Flec_cost%>">
<input type="hidden" name="buying_cost" value="<%=olecture.FOneItem.Fbuying_cost%>">
<input type="hidden" name="matinclude_yn" value="<%=olecture.FOneItem.Fmatinclude_yn%>">
<input type="hidden" name="mat_cost" value="<%=olecture.FOneItem.Fmat_cost%>">
<input type="hidden" name="mat_buying_cost" value="<%=olecture.FOneItem.Fmat_buying_cost%>">

<% if (olecture.FOneItem.Fmatinclude_yn = "C") then %>
	<input type="hidden" name="sellprice" value="<%= (olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost) %>">
	<input type="hidden" name="buycash" value="<%= (olecture.FOneItem.Fbuying_cost + olecture.FOneItem.Fmat_buying_cost) %>">
	<input type="hidden" name="itemsubtotalsum" value="<%= (olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost) %>">
<% else %>
	<input type="hidden" name="sellprice" value="<%= (olecture.FOneItem.Flec_cost) %>">
	<input type="hidden" name="buycash" value="<%=olecture.FOneItem.Fbuying_cost%>">
	<input type="hidden" name="itemsubtotalsum" value="<%=olecture.FOneItem.Flec_cost%>">
<% end if %>

<input type="hidden" name="mileage" value="<%=olecture.FOneItem.Fmileage%>">
<input type="hidden" name="makerId" value="<%=olecture.FOneItem.Flecturer_id%>">
<input type="hidden" name="sitename" value="academy">
<input type="hidden" name="buy_level" value="0">
<input type="hidden" name="weclassyn" value="<%= CHKIIF(olecture.FOneItem.isWeClass,"Y","N") %>">
<tr bgcolor="ffffff">
	<td valign="top" colspan=5>���� : <b><%= lec_idx & " / " & olecture.FOneItem.Flec_title%></b>
	<% if (olecture.FOneItem.isWeClass) then %>
	<b><font color=red>[��ü����]</font></b>
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2">
		<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a">
		<tr align="center" bgcolor="#F7F7F7">
			<td>������<br>����</td>
			<td>���ϸ���</td>
			<td>��û�ο�</td>
			<td>�ѱݾ�<br>����</td>
		</tr>
		<tr align="center" bgcolor="#FFFFFF">
			<td><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %><br><%= FormatNumber(olecture.FOneItem.Fmat_cost,0) %></td>
			<td><%= olecture.FOneItem.Fmileage %> (point)</td>
			<td>
			    <% IF (olecture.FOneItem.isWeClass) THEN %>
			    <input type="text" name="itemea" value="1" size=3 maxlength=3 onChange="chgEntryDetail(this)">
			    <% ELSE %>
				<select name="itemea" onChange="ShowEntryDetail(this)">
					<option value="1">1 ��</option>
					<option value="2">2 ��</option>
					<option value="3">3 ��</option>
					<option value="4">4 ��</option>
				</select>
				<% end if %>
			</td>
			<td id="htmlttlsum">
				<% if (olecture.FOneItem.Fmatinclude_yn = "C") then %>
				<b><%= FormatNumber((olecture.FOneItem.Flec_cost + olecture.FOneItem.Fmat_cost),0) %></b><br>
				����
				<% else %>
				<b><%= FormatNumber(olecture.FOneItem.Flec_cost,0) %></b><br>
				����
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=120 bgcolor="#DDDDFF" align="center">���½ð� ����</td>
	<td><%= getLecOptionBoxHTML(lec_idx,"lecOption","") %></td>
</tr>
<%	For i=0 to 3 %>
<tr id="entry<%=i%>" <% if ((i>0) ) then Response.Write "style='display:none;'"%> bgcolor="#FFFFFF">
	<td width=120 bgcolor="#DDDDFF" align="center">
		������#<%=i+1%>
		<% if i=0 then Response.Write "<br>(�ֹ���)" %>
	</td>
	<td>
		<% if i=0 then %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td width="60" bgcolor="#F8F8F8">���̵�</td>
			<td>
				<input type="text" name="buy_userid" value="" size="12" maxlength="32" class="input" readonly>
				<img src="/images/icon_search.gif" onClick="popSrcId()" style="cursor:pointer" align="absmiddle">
			</td>
		</tr>
		<tr>
			<td width="60" bgcolor="#F8F8F8"><font color="orange">* </font>�� ��</td>
			<td><input type="text" name="buy_name" value="" size="10" maxlength="16" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>��ȭ��ȣ</td>
			<td>
				<input name="buy_phone1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_phone2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_phone3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>�޴���</td>
			<td>
				<input name="buy_hp1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_hp2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="buy_hp3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>�̸���</td>
			<td><input name="buy_email" type="text" size="26" value="" maxlength="90" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>���ϸ��� ��������</td>
			<td><input type="radio" name="mileagegubun" value="ON" checked>���� <input type="radio" name="mileagegubun" value="OFF">��������
			</td>
		</tr>
		</table>
		<% else %>
		<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a">
		<tr>
			<td width="60" bgcolor="#F8F8F8"><font color="orange">* </font>�� ��</td>
			<td><input type="text" name="entryname<%=i%>" value="" size="8" maxlength="16" class="input"></td>
		</tr>
		<tr>
			<td bgcolor="#F8F8F8"><font color="orange">* </font>����ó</td>
			<td>
				<input name="entry<%=i%>_hp1" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="entry<%=i%>_hp2" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
				-
				<input name="entry<%=i%>_hp3" type="text" size="4" maxlength="4" value="" maxlength="4" class="input">
			</td>
		</tr>
		</table>
		<% end if %>
	</td>
</tr>
<%	next %>
<% if olecture.FOneItem.isWeClass THEN %>
<tr  bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF" align="center">��ü����</td>
    <td >
        <table>
        <tbody>
		<tr>
			<th><span>��ü(��ȣȸ)��</span></th>
			<td>
				<span><input type="text" name="wantstudyName" class="txtBasic tblInput" style="width:200px;" maxlength="100" value="" /></span>
				<span class="lPad0">(��ü, ��ȣȸ Ȥ�� ��ǥ�� ���� �Է����ּ���.)</span>
			</td>
		</tr>
		<tr>
			<th><span>���� �����</span></th>
			<td>
				<span>
					<select name="wantstudyYear" class="select tblInput" >
						<option value="2012">2012</option>
						<option value="2013">2013</option>
						<option value="2014">2014</option>
						<option value="2015">2015</option>
						<option value="2016">2016</option>
						<option value="2017">2017</option>
						<option value="2018">2018</option>
						<option value="2019">2019</option>
						<option value="2020">2020</option>
					</select> ��
					<select name="wantstudyMonth" class="select tblInput" >
						<% For i=1 To 12 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> ��
					<select name="wantstudyDay" class="select tblInput" >
						<% For i=1 To 31 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> ��
				</span>
				<span class="lPad0">
					<select name="wantstudyAmPm" class="select tblInput" >
						<option value="����">����</option>
						<option value="����">����</option>
					</select>
					<select name="wantstudyHour" class="select tblInput" >
						<% For i=1 To 12 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> ��
					<select name="wantstudyMin" class="select tblInput" >
						<% For i=0 To 50 step 10 %>
						<option value="<%=i%>"><%=i%></option>
						<% Next %>
					</select> ��
				</span>
			</td>
		</tr>
		<tr>
			<th><span>�������</span></th>
			<td><span><input type="text" name="wantstudyPlace" class="txtBasic tblInput" style="width:500px;" maxlength="100" value="" /></span></td>
		</tr>
		<tr>
			<th><span>���Ǵ��</span></th>
			<td>
				<span><input name="wantstudyWho" type="radio" class="radio" value="1" /> ���</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="2" /> ��ȣȸ</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="3" /> �л�</span>
				<span><input name="wantstudyWho" type="radio" class="radio" value="0" /> ��Ÿ</span>
			</td>
		</tr>
		</tbody>
        </table>
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="#DDDDFF" align="center">�������</td>
	<td>
	    <% if (olecture.FOneItem.isWeClass) then %>
	    <select name="paymethod">
	    <option value="7">�ֹ�����
	    <option value="900">�����Է�(�����Ϸ�)
	    </select>
	    <% else %>
		�����Է�<br>
		<font color="gray">�� ���� : �Է� �Ϸ�� ���� ���´� [�����Ϸ�] �Դϴ�.</font>
		<input type="hidden" name="paymethod" value="900">
		<% end if %>
	</td>
</tr>

<tr bgcolor="ffffff">
	<td valign="top" colspan=5 align="center">
		<img src="/images/icon_save.gif" onClick="SaveItem()" style="cursor:pointer" align="absbottom"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.close()" style="cursor:pointer" align="absbottom">
	</td>
</tr>
</form>
</table>

<%
	set olecture = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->