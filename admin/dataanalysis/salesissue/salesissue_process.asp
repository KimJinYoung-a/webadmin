<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �����ͺм� �����̽�
' History : 2016.01.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim salesidx, department_id,startdate,enddate,title,comment,reguserid,regdate,isusing, startdatetime, enddatetime
dim strSql, mode, menupos, i, lastuserid
	menupos = getNumeric(requestCheckVar(request("menupos"),10))
	salesidx = getNumeric(requestCheckVar(request("salesidx"),10))
	department_id = getNumeric(requestCheckVar(request("department_id"),10))
	startdate = requestCheckVar(request("startdate"),10)
	enddate = requestCheckVar(request("enddate"),10)
	title = request("title")
	comment = request("comment")
	isusing = requestCheckVar(request("isusing"),1)
	startdatetime = requestCheckVar(request("startdatetime"),8)
	enddatetime = requestCheckVar(request("enddatetime"),8)
	mode = requestCheckVar(request("mode"),32)

lastuserid=session("ssBctId")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if mode="salesissuereg" then
	if department_id="" then
		response.write "�μ��� ������ �ּ���."
		dbget.close()	:	response.end
	end if
	if startdate="" then
		response.write "�������� �Է��� �ּ���."
		dbget.close()	:	response.end
	end if
	if enddate="" then
		response.write "�������� �Է��� �ּ���."
		dbget.close()	:	response.end
	end if
	if title="" then
		response.write "������Ʈ���� �Է��� �ּ���."
		dbget.close()	:	response.end
	end if
	if comment="" then
		response.write "����(����/���)�� �Է��� �ּ���."
		dbget.close()	:	response.end
	end if
	if isusing="" then
		response.write "��뿩�θ� ������ �ּ���."
		dbget.close()	:	response.end
	end if

	'/����
	if salesidx<>"" then
		strSql = "Update db_analyze.dbo.tbl_analysis_salesissue" & vbcrlf
		strSql = strSql & " Set department_id="& trim(department_id) &"" & vbcrlf
		strSql = strSql & " ,startdate='" & trim(startdate) & " " & trim(startdatetime) & "'" & vbcrlf
		strSql = strSql & " ,enddate='" & trim(enddate) & " " & trim(enddatetime) & "'" & vbcrlf
		strSql = strSql & " ,title='" & html2db(trim(title)) & "'" & vbcrlf
		strSql = strSql & " ,comment='" & html2db(trim(comment)) & "'" & vbcrlf
		strSql = strSql & " ,isusing='" & trim(isusing) & "' Where " & vbcrlf
		strSql = strSql & " salesidx='"& trim(salesidx) &"'"

		'response.write strSql & "<br>"
		dbanalget.Execute strSql

	'/�űԵ��
	else
		strSql = "insert into db_analyze.dbo.tbl_analysis_salesissue (" & vbcrlf
		strSql = strSql & " department_id, startdate, enddate, title, comment, reguserid, regdate, isusing) values (" & vbcrlf
		strSql = strSql & " "& trim(department_id) &", '" & trim(startdate) & " " & trim(startdatetime) & "', '" & trim(enddate) & " " & trim(enddatetime) & "'" & vbcrlf
		strSql = strSql & " , '" & html2db(trim(title)) & "', '" & html2db(trim(comment)) & "', '" & trim(lastuserid) & "', getdate()" & vbcrlf
		strSql = strSql & " , '" & trim(isusing) & "'" & vbcrlf
		strSql = strSql & " )"
	
		'response.write strSql & "<br>"
		dbanalget.Execute strSql

	end if

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	response.write "	self.close();"
	'response.write "	location.replace('/admin/dataanalysis/salesissue/salesissue_edit.asp?salesidx="& salesidx &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�����ڰ� �����ϴ�.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->