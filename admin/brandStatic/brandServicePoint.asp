<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����귣�弭������
' History : ������ ����
'			2023.11.16 �ѿ�� ����(����ī�װ� �˻� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/brand/brandClass.asp"-->
<%
dim yyyy, mm, makerID, i, dispCate, arrList, CBrandService
	dispCate = requestCheckvar(request("disp"),16)
	yyyy	= req("yyyy1", Left(Date,4))
	mm		= req("mm1", Mid(Date,6,2))
	makerID = req("makerID", "")

set CBrandService = new CBrandServiceList
	CBrandService.frectyyyy = yyyy
	CBrandService.frectmm = mm
	CBrandService.frectmakerID = makerID
	CBrandService.frectdispCate = dispCate
	CBrandService.fBrandServiceList()

if CBrandService.FtotalCount > 0 then
	arrList=CBrandService.fArrList
end if

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
	       	* ���: &nbsp;<% DrawYMBox yyyy,mm %>
			* �귣��ID: <input type="text" class="text" name="makerID" value="<%=makerID%>">
			* ����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->
 
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �������� ȯ������ : 100-(������ҿ���*10) ---> 50���� ��������
		<Br>* ������� ȯ������ : 100-(�������Ǽ�/�����Ǽ�*5) ---> 50���� ��������
		<Br>* CSŬ���� ȯ������ : 100-((ǰ�����+��ǰ+�±�ȯ��)/�����Ǽ�*5) ---> 50���� ��������
		<Br>* ��ǰ���� ȯ������ : 100-��սð�  ---> 50���� ��������

		<Br><Br>* �������� : 4���� ȯ������ ���(��ǰ���ǰ� �������, 3���� ȯ������ ���)
		<Br>* �� �����Ǽ��� 10�� �̸��� ���, �������� ������ ���ǹ��ҵ�. �ϴ� ��� �����ϰ� ���߿� ����
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
    <td colspan="25">
        �˻���� : <b><%= CBrandService.FtotalCount %></b>
    </td>
</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td rowspan="2" width="120">
		<%If makerID <> "" Then %>
			���
		<%Else %>
			�귣��ID
		<%End If%>
		</td>
		<td rowspan="2">�����Ǽ�<br>(��ü���)</td>
        <td colspan="2">������ҿ���</td>
        <td colspan="2">�������(D+4�̻�)</td>
        <td colspan="4">Ŭ���Ӱ���</td>
        <td colspan="3">��ǰ����</td>
		<td rowspan="2" width="60"><b>��������</b></td>
		<td colspan="3">��ǰ�ı�</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>������<br>�ҿ���</td>
        <td><b>ȯ��<br>����</b></td>
        <td>D+4�̻�<br>���Ǽ�</td>
        <td><b>ȯ��<br>����</b></td>
        <td>���<br>(ǰ��)</td>
        <td>��ǰ<br>(�ҷ�/����۵�)</td>
        <td>�±�ȯ<br>(�ҷ�/����۵�)</td>
        <td><b>ȯ��<br>����</b></td>
        <td>��ǰ���ǰǼ�</td>
        <td>�亯�ҿ�ð�</td>
        <td><b>ȯ��<br>����</b></td>
		<td>�ۼ���</td>
		<td>1���ı�</td>
		<td>���<br>����</td>
	</tr>
<%
Dim servicePoint, servicePointText, pnt1, pnt2, pnt3, pnt4
Dim rowCnt
Dim sRs(14)

If IsArray(arrList) Then 
	rowCnt = UBound(arrList,2) + 1
%>

	<%For i=0 To UBound(arrList,2)%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		' Row �ջ�
		sRs(1) = sRs(1) + CDbl(arrList(1,i))			'�ı��
		sRs(2) = sRs(2) + CDbl(arrList(2,i))			'�ı�����
		sRs(3) = sRs(3) + CDbl(arrList(3,i))			'��ǰ���Ǽ�
		sRs(4) = sRs(4) + CDbl(arrList(4,i))			'�亯�ҿ�ð�
		sRs(5) = sRs(5) + CDbl(arrList(5,i))			'����
		sRs(6) = sRs(6) + CDbl(arrList(6,i))			'���ҿ���
		sRs(7) = sRs(7) + CDbl(arrList(7,i))			'ǰ����
		sRs(8) = sRs(8) + CDbl(arrList(8,i))			'��ǰ��
		sRs(9) = sRs(9) + CDbl(arrList(9,i))			'��ȯ��
		sRs(10) = sRs(10) + CDbl(arrList(10,i))		'���������
		sRs(11) = sRs(11) + CDbl(arrList(11,i))		'1�� �ı��
		sRs(12) = sRs(12) + CDbl(arrList(12,i))		'2�� �ı��
		sRs(13) = sRs(13) + CDbl(arrList(13,i))		'3�� �ı��
		sRs(14) = sRs(14) + CDbl(arrList(14,i))		'4�� �ı��

		' ��� �׸� ����
		If CDbl(arrList(1,i)) > 0 Then
			arrList(2,i) = FormatNumber(CDbl(arrList(2,i)) / CDbl(arrList(1,i)) ,2)
		End If 
		If CDbl(arrList(3,i)) > 0 Then
			arrList(4,i) = FormatNumber(CDbl(arrList(4,i)) / CDbl(arrList(3,i)) ,1)
		End If 
		If CDbl(arrList(5,i)) > 0 Then
			arrList(6,i) = FormatNumber(CDbl(arrList(6,i)) / CDbl(arrList(5,i)) ,2)
		End If 

		' ���� ���� ���� ����
		pnt1 = 0
		pnt2 = 0
		pnt3 = 0
		pnt4 = 0

		' �����Ǽ��� ������
		If arrList(5,i) > 0 Then 
			''�������� ȯ������ : 100-(������ҿ���*10) ---> 50���� ��������
			pnt1 = 100 - CInt(10 * CDbl(arrList(6,i)))
			If pnt1 < 50 Then pnt1 = 50

			''������� ȯ������ : 100-(�������Ǽ�/�����Ǽ�%*5) ---> 50���� ��������
			pnt2 = 100 - CInt(500 * CDbl(arrList(10,i)) / CDbl(arrList(5,i)) )
			If pnt2 < 50 Then pnt2 = 50

			''CSŬ���� ȯ������ : 100-((ǰ�����+��ǰ+�±�ȯ��)/�����Ǽ�%*5) ---> 50���� ��������
			pnt3 = 100 - CInt(500 * CDbl(arrList(7,i)+arrList(8,i)+arrList(9,i)) / CDbl(arrList(5,i)))
			If pnt3 < 50 Then pnt3 = 50
		End If 

		' ��ǰ���ǰǼ��� ������
		If arrList(3,i) > 0 Then 
			''��ǰ���� ȯ������ : 100-��սð�  ---> 50���� ��������
			pnt4 = 100 - CLng(arrList(4,i))  ''CInt => CLng ''2016/04/28
			If pnt4 < 50 Then pnt4 = 50
		End If 

		' ���Ǽ��� ������
		If arrList(5,i) > 0 Then
			' ��ǰ���ǰǼ��� ������
			If arrList(3,i) > 0 Then 
				servicePoint = (pnt1 + pnt2 + pnt3 + pnt4) / 4
			Else
				servicePoint = (pnt1 + pnt2 + pnt3) / 3
			End If 

			servicePointText = FormatNumber(servicePoint,2) & "��"
			sRs(0) = sRs(0) + servicePoint

		Else
			servicePointText = "-"
			rowCnt = rowCnt - 1
		End If 

		If pnt1 = 0 Then pnt1 = "-"
		If pnt2 = 0 Then pnt2 = "-"
		If pnt3 = 0 Then pnt3 = "-"
		If pnt4 = 0 Then pnt4 = "-"
	%>
		<td><%=arrList(0,i)%></td>
		<td><%=arrList(5,i)%></td>
		<td><%=arrList(6,i)%>��</td>
		<td><%=pnt1%></td>
		<td><%=arrList(10,i)%></td>
		<td><%=pnt2%></td>

		<td><%=arrList(7,i)%></td>
		<td><%=arrList(8,i)%></td>
		<td><%=arrList(9,i)%></td>
		<td><%=pnt3%></td>

		<td><%=arrList(3,i)%></td>
		<td><%=arrList(4,i)%>�ð�</td>
		<td><%=pnt4%></td>
    	<td><%=servicePointText%></td>
		<td><%=arrList(1,i)%></td>
		<td><%=arrList(11,i)%></td>
		<td><%=arrList(2,i) %>��</td>
	</tr>
	<%Next%>
    <tr align="center" bgcolor="#FFFFFF">
	<%
		If CDbl(sRs(1)) > 0 Then
			sRs(2) = FormatNumber(CDbl(sRs(2)) / CDbl(sRs(1)) ,2)
		End If 
		If CDbl(sRs(3)) > 0 Then
			sRs(4) = FormatNumber(CDbl(sRs(4)) / CDbl(sRs(3)) ,1)
		End If 
		If CDbl(sRs(5)) > 0 Then
			sRs(6) = FormatNumber(CDbl(sRs(6)) / CDbl(sRs(5)) ,2)
		End If
	%>
    	<td><b>�հ� or ���</b></td>
		<td><b><%=FormatNumber(sRs(5),0)%></b></td>
		<td><b><%=sRs(6)%></b>��</td>
		<td>&nbsp;</td>
		<td><b><%=FormatNumber(sRs(10),0)%></b></td>
		<td>&nbsp;</td>

		<td><b><%=FormatNumber(sRs(7),0)%></b></td>
		<td><b><%=FormatNumber(sRs(8),0)%></b></td>
		<td><b><%=FormatNumber(sRs(9),0)%></b></td>
		<td>&nbsp;</td>

		<td><b><%=FormatNumber(sRs(3),0)%></b></td>
		<td><b><%=sRs(4)%></b>�ð�</td>
		<td>&nbsp;</td>
		<td><b><%=FormatNumber( sRs(0) / rowCnt ,2) %></b>��</td>
		<td><b><%=FormatNumber(sRs(1),0)%></b></td>
		<td><b><%=FormatNumber(sRs(11),0)%></b></td>
		<td><b><%=sRs(2)%></b>��</td>
    </tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="25" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<%
End If 
%>
</table>

<%
set CBrandService = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
