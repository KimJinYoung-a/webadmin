<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'// ���� ����
dim msg, lp, menupos
dim mode
dim CateCd, Cate_Name, Cate_NameEng, isusing, orderno
dim SQL
dim page, CateDiv, searchKey, searchString, param, retURL , code_large


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)
CateCd		= RequestCheckvar(Request("CateCd"),3)
code_large	= RequestCheckvar(Request("code_large"),3)
Cate_Name	= html2db(Request("Cate_Name"))
Cate_NameEng= html2db(Request("Cate_NameEng"))
orderno		= RequestCheckvar(Request("orderno"),10)
isusing		= RequestCheckvar(Request("isusing"),1)
page		= RequestCheckvar(Request("page"),10)
CateDiv		= RequestCheckvar(Request("CateDiv"),16)
searchKey	= RequestCheckvar(Request("searchKey"),16)
searchString = Request("searchString")
if Cate_Name <> "" then
	if checkNotValidHTML(Cate_Name) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if Cate_NameEng <> "" then
	if checkNotValidHTML(Cate_NameEng) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end If
if searchString <> "" then
	if checkNotValidHTML(searchString) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
	response.write "</script>"
	response.End
	end if
end if
param = "&page=" & page & "&CateDiv=" & CateDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

'==============================================================================
'## ���� ����(����) ó��

if CateDiv="" then
	response.write	"<script language='javascript'>" &_
					"	alert('ī�װ� ������ �����ϴ�.');" &_
					"	history.back();" &_
					"</script>"
	dbget.close()	:	response.End
end If

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ �űԵ��
		Select Case CateDiv

			Case "code_large"
				'�ߺ��˻�
				SQL = "Select count(code_large) as cnt From db_academy.dbo.tbl_lec_Cate_large where code_large='" & CateCd & "'"
				rsACADEMYget.Open sql, dbACADEMYget, 1
					if rsACADEMYget("cnt")>0 then
						response.write	"<script language='javascript'>" &_
										"	alert('�ߺ��� �ڵ带 �Է��Ͽ����ϴ�.');" &_
										"	history.back();" &_
										"</script>"
						dbget.close()	:	response.End
					end if
				rsACADEMYget.close
		
				'����
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate_large ( code_large , code_nm , orderno ) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "' " &_
						"  , "& orderno &" )"
		
			Case "code_mid"
				'�ߺ��˻�
				SQL = "Select count(code_mid) as cnt From db_academy.dbo.tbl_lec_Cate_mid where code_large =  '" & code_large & "' and code_mid='" & CateCd & "'"
				rsACADEMYget.Open sql, dbACADEMYget, 1
					if rsACADEMYget("cnt")>0 then
						response.write	"<script language='javascript'>" &_
										"	alert('�ߺ��� �ڵ带 �Է��Ͽ����ϴ�.');" &_
										"	history.back();" &_
										"</script>"
						dbget.close()	:	response.End
					end if
				rsACADEMYget.close
		
				'����
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate_mid ( code_large , code_mid , code_nm , code_nm_eng , orderNo ) " &_
						"	Values " &_
						"	( '" & code_large & "'" &_
						"	, '" & CateCd & "'" &_
						"	, '" & Cate_Name & "' " &_
						"	, '" & Cate_NameEng & "' " &_
						"  , " & orderno & " )"
		End Select

		'���� ó��
		dbACADEMYget.Execute(SQL)

		msg = "�ű� ����Ͽ����ϴ�."

		'���ư� ������
		retURL = "categoryList2012.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ����ó��
		Select Case CateDiv

			Case "code_large"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate_large Set " &_
						"	code_nm = '" & Cate_Name & "'" &_
						" Where code_large = '" & CateCd & "'"
			
			Case "code_mid"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate_mid Set " &_
						"	code_nm = '" & Cate_Name & "'" &_
						"	,code_nm_eng = '" & Cate_NameEng & "'" &_
						"	,orderno = " & orderno &_
						"	,display_yn = '" & isusing & "'" &_
						" Where code_large = '" & code_large & "' and code_mid = '"& CateCd &"' "
		End Select

		'���� ó��
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "categoryList2012.asp?menupos=" & menupos & param

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