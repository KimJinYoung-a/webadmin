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
dim CateCd, Cate_Name, Cate_NameEng, isusing, sortNo
dim SQL
dim page, CateDiv, searchKey, searchString, param, retURL


'// ���� ���� �� ó��
menupos		= RequestCheckvar(Request("menupos"),10)
mode		= RequestCheckvar(Request("mode"),16)
CateCd		= RequestCheckvar(Request("CateCd"),3)
Cate_Name	= html2db(Request("Cate_Name"))
Cate_NameEng= html2db(Request("Cate_NameEng"))
sortNo		= RequestCheckvar(Request("sortNo"),10)
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
end if
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
end if

'Ʈ������ ����
dbACADEMYget.beginTrans

Select Case mode
	Case "write"
		'@@ �űԵ��

		Select Case CateDiv
			Case "CateCD1"
				'�ߺ��˻�
				SQL = "Select count(CateCd1) as cnt From db_academy.dbo.tbl_lec_Cate1 where CateCd1='" & CateCd & "'"
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
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate1 (CateCd1, CateCd1_Name) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "') "
			Case "CateCD2"
				'�ߺ��˻�
				SQL = "Select count(CateCd2) as cnt From db_academy.dbo.tbl_lec_Cate2 where CateCd2='" & CateCd & "'"
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
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate2 (CateCd2, CateCd2_Name, CateCd2_Name_Eng) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "'" &_
						"	, '" & Cate_NameEng & "') "
			Case "CateCD3"
				'�ߺ��˻�
				SQL = "Select count(CateCd3) as cnt From db_academy.dbo.tbl_lec_Cate3 where CateCd3='" & CateCd & "'"
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
				SQL =	"Insert into db_academy.dbo.tbl_lec_Cate3 (CateCd3, CateCd3_Name) " &_
						"	Values " &_
						"	( '" & CateCd & "'" &_
						"	, '" & Cate_Name & "') "
		End Select

		'���� ó��
		dbACADEMYget.Execute(SQL)

		msg = "�ű� ����Ͽ����ϴ�."

		'���ư� ������
		retURL = "categoryList.asp?menupos=" & menupos & param

	Case "modify"
		'@@ ����ó��
		Select Case CateDiv
			Case "CateCD1"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate1 Set " &_
						"	CateCd1_Name = '" & Cate_Name & "'" &_
						" Where CateCd1 = '" & CateCd & "'"
			Case "CateCD2"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate2 Set " &_
						"	CateCd2_Name = '" & Cate_Name & "'" &_
						"	,CateCd2_Name_Eng = '" & Cate_NameEng & "'" &_
						"	,sortNo = " & sortNo &_
						"	,isUsing = '" & isusing & "'" &_
						" Where CateCd2 = '" & CateCd & "'"
			Case "CateCD3"
				SQL =	"Update db_academy.dbo.tbl_lec_Cate3 Set " &_
						"	CateCd3_Name = '" & Cate_Name & "'" &_
						"	,sortNo = " & sortNo &_
						"	,isUsing = '" & isusing & "'" &_
						" Where CateCd3 = '" & CateCd & "'"
		End Select

		'���� ó��
		dbACADEMYget.Execute(SQL)

		msg = "�����Ͽ����ϴ�."

		'���ư� ������
		retURL = "categoryList.asp?menupos=" & menupos & param

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