<%@ language=vbscript %>
<% option explicit %>

<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- ������ ����� ������ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" background="/images/menubar_1px.gif">
						<font color="#333333"><b>���ݰ�꼭 ������ �ȳ�</></font>
					</td>
					<td align="right" style="border-bottom:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
						&nbsp;
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    ���1. ���� ���ϰ� (�� bill36524) ���� ����
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
			<img src="/images/icon_num01.gif" border="0"> <b><font color="red">���������� �غ�</font></b><br>
				����ڹ��� ������ Ȥ�� ���ڼ��ݰ�꼭�� �������� �غ��Ͻø� �˴ϴ�.<br>
				1.����ڹ��� ������ : ��� ���ڻ�ŷ��� �̿밡���� �������̸�, ���������� <b>110,000��/��</b> �Դϴ�.<br>
				2.���ݰ�꼭�� ������ : �ŷ����� �湮�ϼż� <b>[���ݰ�꼭�� ������]</b>�� ���� ������ �� �ֽ��ϴ�. �߻������� ���������� <b>��4,400��/��</b> �Դϴ�.<br>
				3.���ϰ� ���� ������ : ���ϰ� �� ����û e���ο����� �̿밡���� �������̸�, ���������� <b>11,000��/��</b> �Դϴ�.<br>
				
				
				<font color="purple">*������ ������������ ������ ��� ����ڴ� �� �������� ����Ͽ� ȸ������ �� ���ݰ�꼭 ������ �����մϴ�.</font><br>
				<font color="purple">*������������ �����ź��� ������ �ŷ����� �湮�ϼż� [���ݰ�꼭�� ������]�� ������ ����Ͻñ� �ٶ��ϴ�.</font><br>
				* ���ϰ� ���� �������� <strong>��õ ���� �ʽ��ϴ�.</strong>(���� ��� ��θ�, ���� �Ⱓ ���� �ɸ�. Ÿ ���ڰ�꼭 ��ü���� ���Ұ�)<br>
				<br>
			
			
			<img src="/images/icon_num02.gif" border="0"> <b><font color="red">���ϰ� ȸ������</font></b><br>
				������������ �غ�Ǽ�����, ���ϰ� �� ȸ�������� �Ͻø� �˴ϴ�.(<a href="https://www.wehago.com" target="_blank">https://www.wehago.com</a>)<br>
				ȸ�����Խÿ��� <b>�����(����/����)ȸ��</b>���� �����Ͻø� �˴ϴ�.<br>
				<font color="purple">ȸ�����Խÿ� ������������ ����� Ȯ���� �����ϹǷ�, ������������ ���� �غ��Ͻñ� �ٶ��ϴ�.</font><br>
				<br>
				
			
			<img src="/images/icon_num03.gif" border="0"> <b><font color="red">�α��� ��, ������ ���</font></b><br>
				ȸ������ �Ϸ� ��, �α��� �Ͻø� ���� ���θ޴��� <b>[�����ȯ�漳��]</b>�̶�� �������� �ֽ��ϴ�.<br>
				[�����ȯ�漳��]���� 4��° �׸� �ִ� <b>������ ���</b>�� ���ֽñ� �ٶ��ϴ�.<br>
				<font color="purple">������ ����� �ȵǾ� ���� ���, �ٹ�����SCM���� ���������� ���� �ʽ��ϴ�.</font><br>
				<br>
				
			
			<img src="/images/icon_num04.gif" border="0"> <b><font color="red">��������� ����Ʈ ����</font></b><br>
				�α��� ������ ǥ�õǴ� ������ ��ܿ� ���ø�, [����]��ư�� �ֽ��ϴ�.<br>
				���ڼ��ݰ�꼭�� ���, �����ڰ� ��������Ḧ �����ϰ� �˴ϴ�. �Ǵ� 200���� ��������ᰡ �ΰ��˴ϴ�.<br>
				�������, 1������ �����Ͻø�, 50���� ���ݰ�꼭 ������ �����մϴ�.<br>
				<br>
				
			
			<img src="/images/icon_num05.gif" border="0"> <b><font color="red">�ٹ�����SCM���� ���ݰ�꼭 ����</font></b><br>
				�� 4���� ������ ��� �غ�Ǿ��ٸ�, �ٹ�����SCM(<a href="https://scm.10x10.co.kr" target="_blank">scm.10x10.co.kr</a>) ���� ���ݰ�꼭�� �����Ͻø� �˴ϴ�.<br>
				<br>

		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    ���2. ����û �̼��� �Ǵ� ��ü �����ü�̿�
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		    1. �̼��� �Ǵ� Ÿ �����ü�� �̿��Ұ�쵵 �������� �ʿ��մϴ�. ����ڹ��������� �Ǵ� ���ݰ�꼭�� �������� �߱޹�������. <br><br>
		    2. ���޹޴��� ���� (�ٹ����� ����ڵ���� <a href="http://scm.10x10.co.kr/images/10x10lic.jpg" target="_blank"><font color="blue">[����]</font></a>) <br><br>
		    3. �� ����� = ���޾�+���� = �հ�ݾ� (�� ����װ� ��꼭 ����ݾ� �հ谡 ��ġ�ؾ� ����Ȯ���˴ϴ�.)<br><br>
		    4. �ۼ�����(������) : <b>�ش� ����� ����</b>(ex] 2013��1������ : 2013-01-31), (���� ������ ������� �����Ͻô´� 1�� ex] 2013��1������ : <%= LEFT(now(),7)&"-01" %> (���ϱ���))<br><br>
		    5. �������ֽ� �̸��� �ּ� : etax@10x10.co.kr<br><br>
			6. ������ �״� 14�ϱ��� ������°� <font color="blue">���� Ȯ��</font>���� �Ǿ� ���� ������� �̸��� ������ ���Ͽ����� ������ ��ȭ �ֽñ� �ٶ��ϴ�.<br>
			
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td colspan="2" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>;border-top:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
	    ���� ���� �Ͻô� ����
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
	    <td width="20" style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-left:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">&nbsp;</td>
		<td style="padding:5; border-bottom:1px solid <%= adminColor("tablebg") %>;border-right:1px solid <%= adminColor("tablebg") %>" bgcolor="#FFFFFF">
		    Q : �� ���ϰ� �� ���������ؾ� �ϳ���? <br>
		    &nbsp;&nbsp; A : �̼��� �Ǵ� ��ü ���� ���α׷����� ���� �� �ּŵ� �˴ϴ�. �ٸ� ���������� �ƴ� ��� �ٹ����ٿ��� ����� ���� ��꼭�� Ȯ���ؾ� �ǹǷ� ���� �� ����Ȯ������ 2~5�� �ҿ�˴ϴ�. <br><br>
		    
		    Q : �����ῡ ���� ���Լ��ݰ�꼭�� ���� �� �ֳ���? <br>
		    &nbsp;&nbsp; A : �� ���⿡�� ������ ���� �ݾ��� �ٹ����ٿ� �����ϴ� ����� ���ϰ� �ֽ��ϴ�. �����ῡ ���� ��꼭�� ���� �������� �ʽ��ϴ�.<br><br>
		    
		    Q : �鼼 ������ε� �� ���ڰ�꼭�� ���� �ؾ� �ǳ���? <br>
		    &nbsp;&nbsp; A : �鼼�� ��� ���� ��꼭 ���� �����մϴ�. ĸ���ؼ� �̸���(etax@10x10.co.kr) �� ���� �����ֽ��� ������ �������� �����ּ���.<br><br>
		    
		    Q : ���ϰ� �� ���� �� ��¹�ư�� ������ �ʽ��ϴ�.<br>
		    &nbsp;&nbsp; A : �鼼�ΰ�� �ٹ����ٿ��� ������ ��� �����ϸ�, �����ΰ�� ����û ������(����) ��� �����մϴ�.<br><br>
		    
		    Q : �������� �����ΰ���?<br>
		    &nbsp;&nbsp; A : ��ü���� ���� > �귣�� ������ ���� �����Ͽ� ���ø� ���� �ֽ��ϴ�(���� �Ϳ� ����). <!-- �̿� �����ϽŰ�� �ۼ����� �Ϳ� 15�ϳ� �����˴ϴ�. -->�������� ��/�Ͽ���,�������� ���, ����(���� ù������)�� �����˴ϴ�. 
		    <br><br>
		    
		    Q : ���ϰ� ��������� ����<br>
		    &nbsp;&nbsp; 1. ���� ����� �ڵ��� ��ȣ�� �ùٸ��� �ʽ��ϴ�. ��ü������������ �������� �ڵ����� 000-000-0000 ��� ���·� ������ ����ϼ���.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; ��ü���� ���� > �������� �ڵ���, �̸��� ������ ������ �ֽñ� �ٶ��ϴ�.<br><br>
		    
		    &nbsp;&nbsp; 2. ���ϰ� ����Ʈ�� ���Ե� ����ڹ�ȣ�� �ٹ����ٿ� ��ϵ� ����ڹ�ȣ�� ��ġ���� �ʽ��ϴ�. ���ϰ� �� ��ϵ� ����ڹ�ȣ: 000-00-00000 �ٹ����ٿ� ��ϵ� ����ڹ�ȣ:XXX-XX-XXXXX<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; �ٹ����ٿ� ��ϵ� ����� ��ȣ�� ���ϰ� �� ��ϵ� ����� ��ȣ�� ��ġ�ؾ� ���� ���� �˴ϴ�. ����ڹ�ȣ�� ����Ȱ�� ����ں��� ��û�� �����մϴ�.(����� ���� ��û�� ��翥�𿡰� ��û�ϼ���. ����ڵ�����纻, ����纻)<br><br>
		    
		    &nbsp;&nbsp; 3. API ����� ���ݰ�꼭<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; ����� �������� ������ ���� ���� �Ǿ�����, ���� Ȯ�����·� ������� ������� �Դϴ�. 1~2���� ����Ȯ�����·� ������� ���������� �����ٶ��ϴ�.  <br><br>
		    
		    &nbsp;&nbsp; 4. ���ϰ� ���� �����ȯ�漳�� => ������ ��Ͽ��� ������ ����� ����Ͻñ� �ٶ��ϴ�.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; �������� ���� �Ǿ��ų�, ������ ����� ���� ������� �߻��ϴ� �޼����Դϴ�. ���ϰ� �α����� ���� ���θ޴��� �����ȯ�漳�� ��ư Ŭ���� ������ �ǿ��� ������ ����� �ٽ� �õ��� �ּ���. <br><br>
		     <img src="/images/Snap_bill_set1.jpg" width="560">
		     <img src="/images/Snap_bill_set2.jpg" width="560">
		    <br><br>
		    &nbsp;&nbsp; 5. ����Ʈ�� �����մϴ�.<br>
		    &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;  -&gt; ���ϰ� �α��� �� ��� [����] Ŭ�� �� ������ ����Ͻñ� �ٶ��ϴ�. �Ǵ� 200�� ������ �߻� <br><br>
		    <img src="/images/Snap_bill_charge1.jpg" width="560">
		    <img src="/images/Snap_bill_charge2.jpg" width="560">
		    <br>
		    
		</td>
	</tr>
</table>
<!-- ������ ����� ���� -->
<!-- #include virtual="/designer/lib/poptail.asp"-->