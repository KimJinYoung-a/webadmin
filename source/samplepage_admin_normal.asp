<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%

dim sellyn,usingyn

%>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :
			&nbsp;
			ī�װ� :
			<br>
			��ǰ�ڵ� :
			<input type="text" class="text" name="" value="" size="32"> (��ǥ�� �����Է°���)
			&nbsp;
			��ǰ�� :
			<input type="text" class="text" name="" value="" size="32" maxlength="32">
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
	     	���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
	     	&nbsp;
	     	����
			<select class="select" name="">
	     	<option value='' selected>��ü</option>
	     	<option value=''>����</option>
	     	<option value=''>MDǰ��</option>
	     	<option value=''>�Ͻ�ǰ��</option>
	     	<option value=''>�����ƴ�</option>
	     	</select>
	     	&nbsp;
	     	����
			<select class="select" name="">
	     	<option value='' selected>��ü</option>
	     	<option value=''>������</option>
	     	<option value=''>����</option>
	     	<option value=''>����(0)</option>
	     	</select>
	     	&nbsp;
	     	�ŷ�����:<% drawSelectBoxMWU "usingyn", usingyn %>
	     	&nbsp;
	     	����
			<select class="select" name="">
	     	<option value='' selected>��ü</option>
	     	<option value=''>����</option>
	     	<option value=''>�鼼</option>
	     	</select>
	     	&nbsp;
	     	����
			<select class="select" name="">
	     	<option value='' selected>��ü</option>
	     	<option value=''>����</option>
	     	<option value=''>���ξ���</option>
	     	</select>
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="��ü����" onClick="">
			&nbsp;
			������ : <input type="text" class="text" name="" size="3" maxlength="5">
			<input type="button" class="button" value="���û�ǰ����" onClick="">
			&nbsp;
			<input type="button" class="button" value="����" onClick="">
			���� ó�� �� �׼��� ���� ���Դϴ�.
		</td>
		<td align="right">
			<img src="/images/icon_star.gif" border="0">
			<img src="/images/icon_plus.gif" border="0">
			<img src="/images/icon_minus.gif" border="0">
			<img src="/images/icon_arrow_up.gif" border="0">
			<img src="/images/icon_arrow_down.gif" border="0">
			<img src="/images/icon_arrow_left.gif" border="0">
			<img src="/images/icon_arrow_right.gif" border="0">

			<img src="/images/question.gif" border="0">

			<img src="/images/btn_word.gif" border="0">
			<img src="/images/btn_excel.gif" border="0">
			<img src="/images/icon_word.gif" border="0">
			<img src="/images/icon_excel.gif" border="0">
			<img src="/images/icon_reload.gif" border="0">
			<img src="/images/icon_go.gif" border="0">


		</td>
	</tr>
</table>
<!-- �׼� �� -->



<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b>350</b>
			&nbsp;
			������ : <b>1 / 20</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="150">adminColor("tablebar")</td>
    	<td width="100"><%= adminColor("tabletop") %></td>
      	<td width="100">�׸�3</td>
      	<td width="100">�׸�4</td>
      	<td width="100">�׸�5</td>
      	<td>���</td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tablebg") %>">
    	<td>adminColor("tablebg")</td>
    	<td><%= adminColor("tablebg") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("pink") %>">
    	<td>adminColor("pink")</td>
    	<td><%= adminColor("pink") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("green") %>">
    	<td>adminColor("green")</td>
    	<td><%= adminColor("green") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("sky") %>">
    	<td>adminColor("sky")</td>
    	<td><%= adminColor("sky") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("gray") %>">
    	<td>adminColor("gray")</td>
    	<td><%= adminColor("gray") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("dgray") %>">
    	<td>adminColor("dgray")</td>
    	<td><%= adminColor("dgray") %></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>&nbsp;</td>
    	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			[pre]
			<font color="red">1</font>
			2
			3
			4
			[next]
		</td>
	</tr>
</table>

<p>

<!-- ���� ���� ���� -->
>>���� CSS
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="150">����</td>
    	<td width="300">���÷���</td>
    	<td width="100">class</td>
      	<td>���</td>

    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>������</td>
    	<td>
    		<input type="button" class="icon" value="#" onClick="">
    		<input type="button" class="icon" value="*" onClick="">
    		<input type="button" class="icon" value="1" onClick="">
    		<input type="button" class="icon" value="2" onClick="">
    		<input type="button" class="icon" value="@" onClick="">
    		<input type="button" class="icon" value=">" onClick="">
    		<input type="button" class="icon" value="711" onClick="">
    	</td>
      	<td>icon</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�Ϲݹ�ư</td>
    	<td><input type="button" class="button" value="��ư" onClick=""></td>
      	<td>button</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�˻���ư</td>
    	<td><input type="button" class="button_s" value="�˻�" onClick=""></td>
      	<td>button_s</td>
      	<td>���Ŀ� Į�� ���� ����(�Ϲݹ�ư�� ����)</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>�����ڹ�ư</td>
    	<td><input type="button" class="button_auth" value="�׼�" onClick=""  ></td>
      	<td>button_s</td>
      	<td>�����ڸ� ���̴� ��ư�Դϴ�.</td>
    </tr>

    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_text</td>
    	<td><input type="text" class="text" name="" ></td>
      	<td>text</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_text (readonly)</td>
    	<td><input type="text" class="text_ro" name="" readonly></td>
      	<td>text_ro</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_textarea</td>
    	<td><textarea class="textarea" name="" rows="2"></textarea></td>
      	<td>textarea</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>input_textarea (readonly)</td>
    	<td><textarea class="textarea_ro" name="" rows="2" readonly></textarea></td>
      	<td>textarea_ro</td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>select</td>
    	<td>
    		<select class="select" name="">
    			<option value='' selected>��ü</option>
	     		<option value=''>�ɼ�01</option>
	     		<option value=''>�ɼ�02</option>
	     		<option value=''>�ɼ�03</option>
	     		<option value=''>�ɼ�04</option>
	     	</select>
    	</td>
      	<td>select</td>
      	<td></td>
    </tr>

</table>




<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->