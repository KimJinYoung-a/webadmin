<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, intloop, menupos
dim mode
dim code_mid, code_large , moveidx , arrmoveidx
dim SQL , retURL
Dim yyyy1, mm1


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)

code_mid	= RequestCheckvar(Request("code_mid"),3)
code_large	= RequestCheckvar(Request("code_large"),3)

moveidx		=	RequestCheckvar(request("moveidx"),10)
yyyy1		=	RequestCheckvar(request("yyyy1"),4)
mm1			=	RequestCheckvar(request("mm1"),2)

If InStr(moveidx,",") > 0 then
	arrmoveidx = split(moveidx,",")
Else  
	arrmoveidx = moveidx
End If 

'==============================================================================
'## ���� ����(����) ó��

if code_large="" then
	response.write	"<script language='javascript'>" &_
					"	alert('ī�װ� ������ �����ϴ�.');" &_
					"	history.back();" &_
					"</script>"
	dbget.close()	:	response.End
end If

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode

	Case "modify"
		'@@ ����ó��

		If InStr(moveidx,",") > 0 then

			For intLoop = 0 To UBound(arrmoveidx)	
					SQL =	"Update db_academy.dbo.tbl_lec_item Set " &_
							"	newCate_Large = '" & code_large & "'" &_
							"	,newCate_mid = '" & code_mid & "'" &_
							" Where idx = '" & Trim(arrmoveidx(intloop)) & "'"
			
				'���� ó��
				dbACADEMYget.Execute(SQL)
			Next 
		
		Else
			
			SQL =	"Update db_academy.dbo.tbl_lec_item Set " &_
							"	newCate_Large = '" & code_large & "'" &_
							"	,newCate_mid = '" & code_mid & "'" &_
							" Where idx = '" & Trim(arrmoveidx) & "'"
			
			'���� ó��
			dbACADEMYget.Execute(SQL)
		
		End If 

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "pop_chg_lec.asp?menupos=" & menupos & "&yyyy1=" & yyyy1 & "&mm1=" & mm1

End Select


'�����˻� �� �ݿ�
If Err.Number = 0 Then   
	dbACADEMYget.CommitTrans				'Ŀ��(����)

	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"	self.location='" & retURL & "';" &_
					"</script>"
Else
    dbACADEMYget.RollBackTrans				'�ѹ�(�����߻���)

	response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.back();" &_
					"</script>"

End If

%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->