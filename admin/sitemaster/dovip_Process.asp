<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : dovip_Process.asp
' Discription : ���ȸ�������ڳ� ó��������
' History : 2015.04.20 ������ ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, sqlStr, lp
dim evt_code, img1, img2, orderby
Dim orgsailprice, orgsailsuplycash, orgsailyn, isusing, idx

menupos		= Request("menupos")
mode		= Request("mode")

evt_code	= Request("evt_code")
orderby	= Request("orderby")
img1		= Request("image1")
img2		= Request("image2")
isusing	= Request("isusing")
idx	= Request("idx")

If isusing="" Then
	isusing="Y"
End If

If orderby="" Then
	orderby="99"
End If

'// Ʈ������ ����
dbget.beginTrans

'// ��忡 ���� �б�
Select Case mode
	Case "add"

		'������ ����
		sqlStr = "Insert Into db_sitemaster.dbo.tbl_vipcorner " &_
				" (evt_code,pcimg,maing,orderby,isusing,regname,regdate) values " &_
				" ('" & evt_code & "'" &_
				" ,'" & img1 &_
				"' ,'" & img2 &_
				"' ,'" & orderby &_
				"' ,'" & isusing &_
				"' ,'" & session("ssBctId") & "'" &_
				" ,getdate())"
				
		dbget.Execute(sqlStr)

	Case "edit"
		'// ���� ����
		sqlStr = "Update [db_sitemaster].[dbo].tbl_vipcorner " &_
				" Set evt_code='" & evt_code &_
				"' 	,pcimg='" & img1 &_
				"' 	,maing='" & img2 &_
				"' 	,orderby='" & orderby & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" 	,modname='" & session("ssBctId") & "'" &_
				" 	,modifydate=getdate()" &_
				" Where idx='" & idx & "'"
		dbget.Execute(sqlStr)
	Case "delete"
		'// ����
		sqlStr = sqlStr & "delete [db_sitemaster].[dbo].tbl_vipcorner " &_
				" Where idx='" & idx & "';" & vbCrLf
		dbget.Execute(sqlStr)

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
	<% if mode="delete" then %>
		alert("�����߽��ϴ�.");
	<% else %>
		alert("�����߽��ϴ�.");
	<% end if %>
	opener.location.reload();
	self.close();
//	self.location = "vip.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->