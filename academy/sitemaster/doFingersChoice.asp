<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
Dim mode,lec_idx,sortNo,MenuId
Dim viewidx,disptitle,allusing
Dim i, ckidx

mode = RequestCheckvar(request("mode"),16)
MenuId = RequestCheckvar(request("MenuId"),10)
lec_idx = request("lec_idx")
sortNo = request("sortNo")
viewidx = request("viewidx")
disptitle = request("disptitle")
allusing = RequestCheckvar(request("allusing"),1)
ckidx= request("ckidx")
  	if lec_idx <> "" then
		if checkNotValidHTML(lec_idx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if sortNo <> "" then
		if checkNotValidHTML(sortNo) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
	if ckidx <> "" then
		if checkNotValidHTML(ckidx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
'// ���۵� ������ �ڵ尪 Ȯ��
If Right(lec_idx,1)="," Then
	lec_idx = Left(lec_idx,Len(lec_idx)-1)
End If

If Right(ckidx,1)="," then
	ckidx = Left(ckidx,Len(ckidx)-1)
End If

Dim sqlStr,msg
on error resume  next 
dbACADEMYget.BeginTrans

If mode="del" Then
	sqlStr = "delete from [db_academy].dbo.tbl_fingersChoice" &_
				" where  idx in (" + ckidx + ") "
ElseIf mode="add" Then
	sqlStr = "insert into [db_academy].dbo.tbl_fingersChoice" &_
				" (MenuId, lec_idx)" &_
				" select '" + Cstr(MenuId) + "', idx" &_
				" from [db_academy].dbo.tbl_lec_item" &_
				" where idx in (" + lec_idx + ")" 

ElseIf mode="isUsingValue" Then
	sqlStr = " update [db_academy].dbo.tbl_fingersChoice " &_
				" set isusing='" & allusing & "'" &_
				" where idx in (" & ckidx & ") "

ElseIf mode="ChangeSort" Then
	lec_idx = split(lec_idx,",")
	sortNo = split(sortNo,",")
	ckidx = split(ckidx,",")
	sqlStr = ""
	For i=0 to ubound(lec_idx)
		sqlStr = sqlStr & " update [db_academy].dbo.tbl_fingersChoice " &_
					" set sortNo='" & sortNo(i) & "'" &_
					" where idx='" & ckidx(i) & "';" & vbCrLf
	Next
End If

dbACADEMYget.execute(sqlStr)

If err.number<>0 Then
	dbACADEMYget.rollback
	msg ="���� �߻�, �����ڹ��� ���"
Else
	dbACADEMYget.committrans
	msg ="���� �Ǿ����ϴ�."
End If
Dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('<%= msg %>');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyClose.asp" -->