<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim id, password, company_name, email, part_sn, posit_sn, level_sn, userdiv, isUsing, mode
	Dim SQL, strMsg

	id			= Request("id")
	password	= Request("password")
	company_name = Request("company_name")
	email		= Request("email")
	part_sn		= Request("part_sn")
	posit_sn	= Request("posit_sn")
	level_sn	= Request("level_sn")
	userdiv		= Request("userdiv")
	isUsing		= Request("isUsing")
	mode		= Request("mode")


	

	'// ó�� �б� //
	Select Case mode
		Case "add"
			strMsg = "���� ������ ��ϵǾ����ϴ�."
			SQL =	"Insert into db_partner.[dbo].tbl_partner " &_
					" (id, password, company_name, email, level_sn, userdiv) values " &_
					" ('" & id & "'" &_
					" ,'" & password & "'" &_				
					" ," & level_sn &_
					" ," & userdiv & ")"					
			dbget.Execute(SQL)
		Case "modi"	
			strMsg = "���� ������ �����Ǿ����ϴ�."
			if userdiv <= 9 then
				dbget.beginTrans
				
				SQL =	"Update db_partner.[dbo].tbl_user_tenbyten Set " &_
					"	part_sn	= " & part_sn &_
					"	,posit_sn	= " & posit_sn &_					
					"Where userid='" & id & "'"
				dbget.Execute(SQL)
			
				SQL =	"Update db_partner.[dbo].tbl_partner Set lastInfoChgDT=getdate(), " &_
					"	password	= '" & password & "' " &_					
					"	,level_sn	= " & level_sn &_
					"	,userdiv	= '" & userdiv & "' " &_
					"Where id='" & id & "'"
			dbget.Execute(SQL)
			
				If Err.Number = 0 Then
					dbget.CommitTrans	
				else
				 	dbget.RollBackTrans	
				end if
			else	
				SQL =	"Update db_partner.[dbo].tbl_partner Set lastInfoChgDT=getdate(), " &_
						"	password	= '" & password & "' " &_
						"	,level_sn	= " & level_sn &_
						"	,userdiv	= '" & userdiv & "' " &_
						"Where id='" & id & "'"
				dbget.Execute(SQL)
			end if
		Case "del"
			strMsg = "ó���� �Ϸ�Ǿ����ϴ�."
			SQL =	"Update db_partner.[dbo].tbl_partner Set lastInfoChgDT=getdate(), " &_
					"	isUsing = '" + isUsing + "' " &_
					"Where id='" & id & "'"
			dbget.Execute(SQL)
			
			SQL = "Update db_partner.dbo.tbl_user_tenbyten Set userid = Null " &_
				"	WHERE userid = '"&id&"'"
			dbget.Execute(SQL)				
	End Select

	'�����˻� �� ����
	If Err.Number = 0 Then   
		response.write	"<script language='javascript'>" &_
						"	alert('" & strMsg & "');" &_
						"	opener.history.go(0);" &_
						"	self.close();" &_
						"</script>"
	Else	
		response.write	"<script language='javascript'>" &_
						"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
						"	history.back();" &_
						"</script>"
	
	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->