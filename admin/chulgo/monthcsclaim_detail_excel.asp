<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  CS������Ŭ�������
' History : 2007.08.22 �ѿ�� ����
'###########################################################
%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/chulgoclass/chulgoclass.asp" -->

<%
dim yyyy
	yyyy = request("yyyy")
		
dim omonthcsclaimtotal , i
	set omonthcsclaimtotal = new Cchulgoitemlist
	omonthcsclaimtotal.frectyyyy = yyyy
	omonthcsclaimtotal.fmonthcsclaimtotal()
%>

<%
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"monthcsclaim_detail_"+yyyy+".xls"
%>
<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="25" valign="top">
		<td>
			<font color="red"><strong> <%= yyyy %> CS ������ Ŭ���� ���</strong></font>
		</td>
	</tr>
</table>
<!--ǥ ��峡-->

<% dim fgong1total,fgong2total,fgong3total,fgong4total,fgong5total,fitem1total,fitem2total,fitem3total,fitem4total,fitem5total
dim fmul1total,fmul2total,fmul3total,fmul4total,fmul5total,fmul6total,fmul7total,ftak1total,ftak2total,ftak3total,fgitatotal
dim fa000total,fa001total,fa002total,fa004total,fa010total,fa011total,fa008total,ftotalsum
%>
<!-- ��������-->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	
	<tr bgcolor=#DDDDFF>
		<td align="center" colspan="2">����</td>
		<td align="center">�±�ȯ���</td>
		<td align="center">������߼�</td>
		<td align="center">���񽺹߼�</td>
		<td align="center">��ǰ</td>
		<td align="center">ȸ��</td>
		<td align="center">�±�ȯȸ��</td>
		<td align="center">�ֹ����</td>
		<td align="center">�հ�</td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center" rowspan="5">����</td>
		<td align="center">�ܼ�����</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gong1 %></td>
		<td align="center"><% fgong1total= omonthcsclaimtotal.flist(i).fa000gong1+omonthcsclaimtotal.flist(i).fa001gong1+omonthcsclaimtotal.flist(i).fa002gong1+omonthcsclaimtotal.flist(i).fa004gong1+omonthcsclaimtotal.flist(i).fa010gong1+omonthcsclaimtotal.flist(i).fa011gong1+omonthcsclaimtotal.flist(i).fa008gong1 %>
		<%=fgong1total%></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">���ֹ�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gong2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gong2 %></td>
		<td align="center"><% fgong2total= omonthcsclaimtotal.flist(i).fa000gong2+omonthcsclaimtotal.flist(i).fa001gong2+omonthcsclaimtotal.flist(i).fa002gong2+omonthcsclaimtotal.flist(i).fa004gong2+omonthcsclaimtotal.flist(i).fa010gong2+omonthcsclaimtotal.flist(i).fa011gong2+omonthcsclaimtotal.flist(i).fa008gong2 %>
		<%=fgong2total%></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">������</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gong3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gong3 %></td>
		<td align="center"><% fgong3total= omonthcsclaimtotal.flist(i).fa000gong3+omonthcsclaimtotal.flist(i).fa001gong3+omonthcsclaimtotal.flist(i).fa002gong3+omonthcsclaimtotal.flist(i).fa004gong3+omonthcsclaimtotal.flist(i).fa010gong3+omonthcsclaimtotal.flist(i).fa011gong3+omonthcsclaimtotal.flist(i).fa008gong3 %>
		<%=fgong3total%></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">ǰ��</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gong4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gong4 %></td>
		<td align="center"><% fgong4total= omonthcsclaimtotal.flist(i).fa000gong4+omonthcsclaimtotal.flist(i).fa001gong4+omonthcsclaimtotal.flist(i).fa002gong4+omonthcsclaimtotal.flist(i).fa004gong4+omonthcsclaimtotal.flist(i).fa010gong4+omonthcsclaimtotal.flist(i).fa011gong4+omonthcsclaimtotal.flist(i).fa008gong4 %>
		<%=fgong4total%></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��Ÿ</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gong1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gong5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gong5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gong5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gong5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gong5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gong5 %></td>
		<td align="center"><% fgong5total= omonthcsclaimtotal.flist(i).fa000gong5+omonthcsclaimtotal.flist(i).fa001gong5+omonthcsclaimtotal.flist(i).fa002gong5+omonthcsclaimtotal.flist(i).fa004gong1+omonthcsclaimtotal.flist(i).fa010gong5+omonthcsclaimtotal.flist(i).fa011gong5+omonthcsclaimtotal.flist(i).fa008gong5 %>
		<%=fgong5total%></td>
	</tr>
	<tr bgcolor=#FFFFFF>
		<td align="center" rowspan="5">��ǰ����</td>
		<td align="center">��ǰ�ҷ�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011item1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008item1 %></td>
		<td align="center"><% fitem1total=omonthcsclaimtotal.flist(i).fa000item1+omonthcsclaimtotal.flist(i).fa001item1+omonthcsclaimtotal.flist(i).fa002item1+omonthcsclaimtotal.flist(i).fa004item1+omonthcsclaimtotal.flist(i).fa010item1+omonthcsclaimtotal.flist(i).fa011item1+omonthcsclaimtotal.flist(i).fa008item1 %>
		<%= fitem1total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��ǰ�Ҹ���</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011item2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008item2 %></td>
		<td align="center"><% fitem2total=omonthcsclaimtotal.flist(i).fa000item2+omonthcsclaimtotal.flist(i).fa001item2+omonthcsclaimtotal.flist(i).fa002item2+omonthcsclaimtotal.flist(i).fa004item2+omonthcsclaimtotal.flist(i).fa010item2+omonthcsclaimtotal.flist(i).fa011item2+omonthcsclaimtotal.flist(i).fa008item2 %>
		<%= fitem2total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��ǰ��Ͽ���</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011item3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008item3 %></td>
		<td align="center"><% fitem3total=omonthcsclaimtotal.flist(i).fa000item3+omonthcsclaimtotal.flist(i).fa001item3+omonthcsclaimtotal.flist(i).fa002item3+omonthcsclaimtotal.flist(i).fa004item3+omonthcsclaimtotal.flist(i).fa010item3+omonthcsclaimtotal.flist(i).fa011item3+omonthcsclaimtotal.flist(i).fa008item3 %>
		<%= fitem3total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��ǰ����ҷ�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011item4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008item4 %></td>
		<td align="center"><% fitem4total=omonthcsclaimtotal.flist(i).fa000item4+omonthcsclaimtotal.flist(i).fa001item4+omonthcsclaimtotal.flist(i).fa002item4+omonthcsclaimtotal.flist(i).fa004item4+omonthcsclaimtotal.flist(i).fa010item4+omonthcsclaimtotal.flist(i).fa011item4+omonthcsclaimtotal.flist(i).fa008item4 %>
		<%= fitem4total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��Ÿ</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011item5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008item5 %></td>
		<td align="center"><% fitem5total=omonthcsclaimtotal.flist(i).fa000item5+omonthcsclaimtotal.flist(i).fa001item5+omonthcsclaimtotal.flist(i).fa002item5+omonthcsclaimtotal.flist(i).fa004item5+omonthcsclaimtotal.flist(i).fa010item5+omonthcsclaimtotal.flist(i).fa011item5+omonthcsclaimtotal.flist(i).fa008item5 %>
		<%= fitem5total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center" rowspan="7">��������</td>
		<td align="center">���߼�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul1 %></td>
		<td align="center"><% fmul1total= omonthcsclaimtotal.flist(i).fa000mul1+omonthcsclaimtotal.flist(i).fa001mul1+omonthcsclaimtotal.flist(i).fa002mul1+omonthcsclaimtotal.flist(i).fa004mul1+omonthcsclaimtotal.flist(i).fa010mul1+omonthcsclaimtotal.flist(i).fa011mul1+omonthcsclaimtotal.flist(i).fa008mul1 %>
		<%= fmul1total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��ǰ�ļ�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul2 %></td>
		<td align="center"><% fmul2total= omonthcsclaimtotal.flist(i).fa000mul2+omonthcsclaimtotal.flist(i).fa001mul2+omonthcsclaimtotal.flist(i).fa002mul2+omonthcsclaimtotal.flist(i).fa004mul2+omonthcsclaimtotal.flist(i).fa010mul2+omonthcsclaimtotal.flist(i).fa011mul2+omonthcsclaimtotal.flist(i).fa008mul2 %>
		<%= fmul2total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">���Ż�ǰ����</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul3 %></td>
		<td align="center"><% fmul3total= omonthcsclaimtotal.flist(i).fa000mul3+omonthcsclaimtotal.flist(i).fa001mul3+omonthcsclaimtotal.flist(i).fa002mul3+omonthcsclaimtotal.flist(i).fa004mul3+omonthcsclaimtotal.flist(i).fa010mul3+omonthcsclaimtotal.flist(i).fa011mul3+omonthcsclaimtotal.flist(i).fa008mul3 %>
		<%= fmul3total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">����ǰ����</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul4 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul4 %></td>
		<td align="center"><% fmul4total= omonthcsclaimtotal.flist(i).fa000mul4+omonthcsclaimtotal.flist(i).fa001mul4+omonthcsclaimtotal.flist(i).fa002mul4+omonthcsclaimtotal.flist(i).fa004mul4+omonthcsclaimtotal.flist(i).fa010mul4+omonthcsclaimtotal.flist(i).fa011mul4+omonthcsclaimtotal.flist(i).fa008mul4 %>
		<%= fmul4total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��ǰǰ��</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul5 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul5 %></td>
		<td align="center"><% fmul5total= omonthcsclaimtotal.flist(i).fa000mul5+omonthcsclaimtotal.flist(i).fa001mul5+omonthcsclaimtotal.flist(i).fa002mul5+omonthcsclaimtotal.flist(i).fa004mul5+omonthcsclaimtotal.flist(i).fa010mul5+omonthcsclaimtotal.flist(i).fa011mul5+omonthcsclaimtotal.flist(i).fa008mul5 %>
		<%= fmul5total %></td>
	</tr>		
	<tr bgcolor=#FFFFFF>
		<td align="center">�������</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul6 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul6 %></td>
		<td align="center"><% fmul6total= omonthcsclaimtotal.flist(i).fa000mul6+omonthcsclaimtotal.flist(i).fa001mul6+omonthcsclaimtotal.flist(i).fa002mul6+omonthcsclaimtotal.flist(i).fa004mul6+omonthcsclaimtotal.flist(i).fa010mul6+omonthcsclaimtotal.flist(i).fa011mul6+omonthcsclaimtotal.flist(i).fa008mul6 %>
		<%= fmul6total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��Ÿ</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011mul7 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008mul7 %></td>
		<td align="center"><% fmul7total= omonthcsclaimtotal.flist(i).fa000mul7+omonthcsclaimtotal.flist(i).fa001mul7+omonthcsclaimtotal.flist(i).fa002mul7+omonthcsclaimtotal.flist(i).fa004mul7+omonthcsclaimtotal.flist(i).fa010mul7+omonthcsclaimtotal.flist(i).fa011mul7+omonthcsclaimtotal.flist(i).fa008mul7 %>
		<%= fmul7total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center" rowspan="3">��������</td>
		<td align="center">�������</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011tak1 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008tak1 %></td>
		<td align="center"><% ftak1total= omonthcsclaimtotal.flist(i).fa000tak1+omonthcsclaimtotal.flist(i).fa001tak1+omonthcsclaimtotal.flist(i).fa002tak1+omonthcsclaimtotal.flist(i).fa004tak1+omonthcsclaimtotal.flist(i).fa010tak1+omonthcsclaimtotal.flist(i).fa011tak1+omonthcsclaimtotal.flist(i).fa008tak1 %>
		<%= ftak1total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">�ù���ļ�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011tak2 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008tak2 %></td>
		<td align="center"><% ftak2total= omonthcsclaimtotal.flist(i).fa000tak2+omonthcsclaimtotal.flist(i).fa001tak2+omonthcsclaimtotal.flist(i).fa002tak2+omonthcsclaimtotal.flist(i).fa004tak2+omonthcsclaimtotal.flist(i).fa010tak2+omonthcsclaimtotal.flist(i).fa011tak2+omonthcsclaimtotal.flist(i).fa008tak2 %>
		<%= ftak2total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">�ù��н�</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011tak3 %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008tak3 %></td>
		<td align="center"><% ftak3total= omonthcsclaimtotal.flist(i).fa000tak3+omonthcsclaimtotal.flist(i).fa001tak3+omonthcsclaimtotal.flist(i).fa002tak3+omonthcsclaimtotal.flist(i).fa004tak3+omonthcsclaimtotal.flist(i).fa010tak3+omonthcsclaimtotal.flist(i).fa011tak3+omonthcsclaimtotal.flist(i).fa008tak3 %>
		<%= ftak3total %></td>
	</tr>	
	<tr bgcolor=#FFFFFF>
		<td align="center">��Ÿ</td>
		<td align="center">��Ÿ</td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa000gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa001gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa002gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa004gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa010gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa011gita %></td>
		<td align="center"><%= omonthcsclaimtotal.flist(i).fa008gita %></td>
		<td align="center"><% fgitatotal= omonthcsclaimtotal.flist(i).fa000gita+omonthcsclaimtotal.flist(i).fa001gita+omonthcsclaimtotal.flist(i).fa002gita+omonthcsclaimtotal.flist(i).fa004gita+omonthcsclaimtotal.flist(i).fa010gita+omonthcsclaimtotal.flist(i).fa011gita+omonthcsclaimtotal.flist(i).fa008gita %>
		<%= fgitatotal %></td>
	</tr>
	<tr bgcolor=#DDDDFF>
		<td align="center" colspan="2">�հ�</td>
		<td align="center"><% fa000total=omonthcsclaimtotal.flist(i).fa000gong1+omonthcsclaimtotal.flist(i).fa000gong2+omonthcsclaimtotal.flist(i).fa000gong3+omonthcsclaimtotal.flist(i).fa000gong4+omonthcsclaimtotal.flist(i).fa000gong5+omonthcsclaimtotal.flist(i).fa000item1+omonthcsclaimtotal.flist(i).fa000item2+omonthcsclaimtotal.flist(i).fa000item3+omonthcsclaimtotal.flist(i).fa000item4+omonthcsclaimtotal.flist(i).fa000item5+omonthcsclaimtotal.flist(i).fa000mul1+omonthcsclaimtotal.flist(i).fa000mul2+omonthcsclaimtotal.flist(i).fa000mul3+omonthcsclaimtotal.flist(i).fa000mul4+omonthcsclaimtotal.flist(i).fa000mul5+omonthcsclaimtotal.flist(i).fa000mul6+omonthcsclaimtotal.flist(i).fa000mul7+omonthcsclaimtotal.flist(i).fa000tak1+omonthcsclaimtotal.flist(i).fa000tak2+omonthcsclaimtotal.flist(i).fa000tak1+omonthcsclaimtotal.flist(i).fa000gita %>
		<%= fa000total %></td>
		<td align="center"><% fa001total = omonthcsclaimtotal.flist(i).fa001gong1+omonthcsclaimtotal.flist(i).fa001gong2+omonthcsclaimtotal.flist(i).fa001gong3+omonthcsclaimtotal.flist(i).fa001gong4+omonthcsclaimtotal.flist(i).fa001gong5+omonthcsclaimtotal.flist(i).fa001item1+omonthcsclaimtotal.flist(i).fa001item2+omonthcsclaimtotal.flist(i).fa001item3+omonthcsclaimtotal.flist(i).fa001item4+omonthcsclaimtotal.flist(i).fa001item5+omonthcsclaimtotal.flist(i).fa001mul1+omonthcsclaimtotal.flist(i).fa001mul2+omonthcsclaimtotal.flist(i).fa001mul3+omonthcsclaimtotal.flist(i).fa001mul4+omonthcsclaimtotal.flist(i).fa001mul5+omonthcsclaimtotal.flist(i).fa001mul6+omonthcsclaimtotal.flist(i).fa001mul7+omonthcsclaimtotal.flist(i).fa001tak1+omonthcsclaimtotal.flist(i).fa001tak2+omonthcsclaimtotal.flist(i).fa001tak3+omonthcsclaimtotal.flist(i).fa001gita %>
		<%= fa001total %>
		</td>
		<td align="center"><% fa002total= omonthcsclaimtotal.flist(i).fa002gong1+omonthcsclaimtotal.flist(i).fa002gong2+omonthcsclaimtotal.flist(i).fa002gong3+omonthcsclaimtotal.flist(i).fa002gong4+omonthcsclaimtotal.flist(i).fa002gong5+omonthcsclaimtotal.flist(i).fa002item1+omonthcsclaimtotal.flist(i).fa002item2+omonthcsclaimtotal.flist(i).fa002item3+omonthcsclaimtotal.flist(i).fa002item4+omonthcsclaimtotal.flist(i).fa002item5+omonthcsclaimtotal.flist(i).fa002mul1+omonthcsclaimtotal.flist(i).fa002mul2+omonthcsclaimtotal.flist(i).fa002mul3+omonthcsclaimtotal.flist(i).fa002mul4+omonthcsclaimtotal.flist(i).fa002mul5+omonthcsclaimtotal.flist(i).fa002mul6+omonthcsclaimtotal.flist(i).fa002mul7+omonthcsclaimtotal.flist(i).fa002tak1+omonthcsclaimtotal.flist(i).fa002tak2+omonthcsclaimtotal.flist(i).fa002tak3+omonthcsclaimtotal.flist(i).fa002gita %>
		<%= fa002total %></td>	
		<td align="center"><% fa004total= omonthcsclaimtotal.flist(i).fa004gong1+omonthcsclaimtotal.flist(i).fa004gong2+omonthcsclaimtotal.flist(i).fa004gong3+omonthcsclaimtotal.flist(i).fa004gong4+omonthcsclaimtotal.flist(i).fa004gong5+omonthcsclaimtotal.flist(i).fa004item1+omonthcsclaimtotal.flist(i).fa004item2+omonthcsclaimtotal.flist(i).fa004item3+omonthcsclaimtotal.flist(i).fa004item4+omonthcsclaimtotal.flist(i).fa004item5+omonthcsclaimtotal.flist(i).fa004mul1+omonthcsclaimtotal.flist(i).fa004mul2+omonthcsclaimtotal.flist(i).fa004mul3+omonthcsclaimtotal.flist(i).fa004mul4+omonthcsclaimtotal.flist(i).fa004mul5+omonthcsclaimtotal.flist(i).fa004mul6+omonthcsclaimtotal.flist(i).fa004mul7+omonthcsclaimtotal.flist(i).fa004tak1+omonthcsclaimtotal.flist(i).fa004tak2+omonthcsclaimtotal.flist(i).fa004tak3+omonthcsclaimtotal.flist(i).fa004gita %>
		<%= fa004total %></td>
		<td align="center"><% fa010total= omonthcsclaimtotal.flist(i).fa010gong1+omonthcsclaimtotal.flist(i).fa010gong2+omonthcsclaimtotal.flist(i).fa010gong3+omonthcsclaimtotal.flist(i).fa010gong4+omonthcsclaimtotal.flist(i).fa010gong5+omonthcsclaimtotal.flist(i).fa010item1+omonthcsclaimtotal.flist(i).fa010item2+omonthcsclaimtotal.flist(i).fa010item3+omonthcsclaimtotal.flist(i).fa010item4+omonthcsclaimtotal.flist(i).fa010item5+omonthcsclaimtotal.flist(i).fa010mul1+omonthcsclaimtotal.flist(i).fa010mul2+omonthcsclaimtotal.flist(i).fa010mul3+omonthcsclaimtotal.flist(i).fa010mul4+omonthcsclaimtotal.flist(i).fa010mul5+omonthcsclaimtotal.flist(i).fa010mul6+omonthcsclaimtotal.flist(i).fa010mul7+omonthcsclaimtotal.flist(i).fa010tak1+omonthcsclaimtotal.flist(i).fa010tak2+omonthcsclaimtotal.flist(i).fa010tak3+omonthcsclaimtotal.flist(i).fa010gita %>
		<%= fa010total %></td>
		<td align="center"><% fa011total= omonthcsclaimtotal.flist(i).fa011gong1+omonthcsclaimtotal.flist(i).fa011gong2+omonthcsclaimtotal.flist(i).fa011gong3+omonthcsclaimtotal.flist(i).fa011gong4+omonthcsclaimtotal.flist(i).fa011gong5+omonthcsclaimtotal.flist(i).fa011item1+omonthcsclaimtotal.flist(i).fa011item2+omonthcsclaimtotal.flist(i).fa011item3+omonthcsclaimtotal.flist(i).fa011item4+omonthcsclaimtotal.flist(i).fa011item5+omonthcsclaimtotal.flist(i).fa011mul1+omonthcsclaimtotal.flist(i).fa011mul2+omonthcsclaimtotal.flist(i).fa011mul3+omonthcsclaimtotal.flist(i).fa011mul4+omonthcsclaimtotal.flist(i).fa011mul5+omonthcsclaimtotal.flist(i).fa011mul6+omonthcsclaimtotal.flist(i).fa011mul7+omonthcsclaimtotal.flist(i).fa011tak1+omonthcsclaimtotal.flist(i).fa011tak2+omonthcsclaimtotal.flist(i).fa011tak3+omonthcsclaimtotal.flist(i).fa011gita %>
		<%= fa011total %></td>
		<td align="center"><% fa008total= omonthcsclaimtotal.flist(i).fa008gong1+omonthcsclaimtotal.flist(i).fa008gong2+omonthcsclaimtotal.flist(i).fa008gong3+omonthcsclaimtotal.flist(i).fa008gong4+omonthcsclaimtotal.flist(i).fa008gong5+omonthcsclaimtotal.flist(i).fa008item1+omonthcsclaimtotal.flist(i).fa008item2+omonthcsclaimtotal.flist(i).fa008item3+omonthcsclaimtotal.flist(i).fa008item4+omonthcsclaimtotal.flist(i).fa008item5+omonthcsclaimtotal.flist(i).fa008mul1+omonthcsclaimtotal.flist(i).fa008mul2+omonthcsclaimtotal.flist(i).fa008mul3+omonthcsclaimtotal.flist(i).fa008mul4+omonthcsclaimtotal.flist(i).fa008mul5+omonthcsclaimtotal.flist(i).fa008mul6+omonthcsclaimtotal.flist(i).fa008mul7+omonthcsclaimtotal.flist(i).fa008tak1+omonthcsclaimtotal.flist(i).fa008tak2+omonthcsclaimtotal.flist(i).fa008tak3+omonthcsclaimtotal.flist(i).fa008gita %>
		<%= fa008total %></td>
		<td align="center"><% ftotalsum= fgong1total+fgong2total+fgong3total+fgong4total+fgong5total+fitem1total+fitem2total+fitem3total+fitem4total+fitem5total+fmul1total+fmul2total+fmul3total+fmul4total+fmul5total+fmul6total+fmul7total+ftak1total+ftak2total+ftak3total+fgitatotal %>
		<%= ftotalsum %></td>
	</tr>
</table>		
<!-- ������-->		


<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
