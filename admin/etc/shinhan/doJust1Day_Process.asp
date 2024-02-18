<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doJust1Day_Process.asp
' Discription : ��������� ����Ʈ ������ ó�� ������
' History : 2009.10.27 ������ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, sqlStr, lp
dim justDate, itemid, salePrice, orgPrice, saleSuplyCash, limitNo, justDesc

menupos		= Request("menupos")
mode		= Request("mode")

justDate	= Request("justDate")
itemid		= Request("itemid")
salePrice	= Request("salePrice")
orgPrice	= Request("orgPrice")
saleSuplyCash = Request("saleSuplyCash")
limitNo		= Request("limitNo")
justDesc	= html2db(Request("justDesc"))

'// Ʈ������ ����
dbget.beginTrans

'// ��忡 ���� �б�
Select Case mode
	Case "add"
		'// �ű� ���
		rsget.Open "Select count(JustDate) from db_temp.[dbo].tbl_just1Day_Shinhan where JustDate='" & justDate & "'", dbget, 1
		if rsget(0)>0 then
			Alert_return("�̹� ��ϵ� ��¥�Դϴ�.\n�ٸ� ��¥�� �������ּ���.")
			dbget.close()	:	response.End
		end if
		rsget.Close

		'����Ʈ ������ ����
		sqlStr = "Insert Into db_temp.[dbo].tbl_just1Day_Shinhan " &_
				" (JustDate,itemid,orgPrice,justSalePrice,SaleSuplyCash,justDesc,limitNo,adminid) values " &_
				" ('" & justDate & "'" &_
				" ," & itemid &_
				" ," & orgPrice &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ,'" & justDesc & "'" &_
				" ," & limitNo &_
				" ,'" & session("ssBctId") & "')"
		dbget.Execute(sqlStr)

	Case "edit"
		'// ���� ����
		sqlStr = "Update db_temp.[dbo].tbl_just1Day_Shinhan " &_
				" Set justSalePrice=" & salePrice &_
				" 	,SaleSuplyCash=" & SaleSuplyCash &_
				" 	,limitNo=" & limitNo &_
				" 	,justDesc='" & justDesc & "'" &_
				" Where justDate='" & justDate & "'"
		dbget.Execute(sqlStr)
        
	Case "delete"
		'// ����
		if justDate>cStr(date()) then
			'����Ʈ������ ���� ����
			sqlStr = sqlStr & "delete db_temp.[dbo].tbl_just1Day_Shinhan " &_
					" Where justDate='" & justDate & "';" & vbCrLf
			dbget.Execute(sqlStr)
		else
			Alert_return("���� �������̰ų� �Ϸ�� ��ǰ�� ������ �� �����ϴ�.")
			dbget.close()	:	response.End
		end if

End Select


'// Ʈ������ �˻� �� ����
If Err.Number = 0 Then
        dbget.CommitTrans
Else
        dbget.RollBackTrans
		Alert_return("����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.")
		dbget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "Just1Day_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->