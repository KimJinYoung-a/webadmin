<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ����FAQ
' Hieditor : 2021.02.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/board/faq/customer_manual_faq_cls.asp"-->
<%
Dim menupos,mode, sql,fidx,gubun,contents,solution,isusing,regdate,lastupdate,lastadminid, manualtype
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
    mode = requestCheckVar(request("mode"),32)
    fidx = requestCheckVar(getNumeric(request("fidx")),10)
    gubun = requestCheckVar(request("gubun"),10)
    contents = requestCheckVar(trim(request("contents")),512)
    solution = trim(request("solution"))
    isusing = requestCheckVar(request("isusing"),1)

manualtype="staff_faq"
lastadminid=session("ssBctId")

if mode = "faqreg" then

	'�űԵ��
	if fidx = "" then
        if contents <> "" and not(isnull(contents)) then
        	contents = ReplaceBracket(contents)

            'if checkNotValidHTML(contents) then
            '    response.write "<script type='text/javascript'>"
            '    response.write "	alert('���ǳ��뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
            '    response.write "</script>"
            '    dbget.close() : response.end
            'end if
        end If
        if solution <> "" and not(isnull(solution)) then
        	solution = ReplaceBracket(solution)

            'if checkNotValidHTML(solution) then
            '    response.write "<script type='text/javascript'>"
            '    response.write "	alert('ó������� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
            '    response.write "</script>"
            '    dbget.close() : response.end
            'end if
        end If

		sql = "INSERT INTO db_cs.dbo.tbl_customer_manual_faq" + vbcrlf
		sql = sql & " (manualtype,gubun,contents,solution,isusing,regdate,lastupdate,lastadminid) values (" + vbcrlf
		sql = sql & " '"& manualtype &"',"& gubun &",'"& html2db(contents) &"','"& html2db(solution) &"','"&isusing&"',getdate(),getdate(),'"& lastadminid &"'" + vbcrlf
		sql = sql & " )"

		'response.write sql &"<br>"
		dbget.execute sql
				
	'//�������	
	else
        if contents <> "" and not(isnull(contents)) then
        	contents = ReplaceBracket(contents)

            'if checkNotValidHTML(contents) then
            '    response.write "<script type='text/javascript'>"
            '    response.write "	alert('���ǳ��뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
            '    response.write "</script>"
            '    dbget.close() : response.end
            'end if
        end If
        if solution <> "" and not(isnull(solution)) then
        	solution = ReplaceBracket(solution)

            'if checkNotValidHTML(solution) then
            '    response.write "<script type='text/javascript'>"
            '    response.write "	alert('ó������� ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
            '    response.write "</script>"
            '    dbget.close() : response.end
            'end if
        end If

		sql = "UPDATE db_cs.dbo.tbl_customer_manual_faq" + vbcrlf
		sql = sql & " SET manualtype='"& manualtype &"'" + vbcrlf
        sql = sql & " ,gubun = "& gubun &"" + vbcrlf
		sql = sql & " ,contents = '"& html2db(contents) &"'" + vbcrlf
		sql = sql & " ,solution = '"& html2db(solution) &"'" + vbcrlf
        sql = sql & " ,isusing = '"& isusing &"'" + vbcrlf
		sql = sql & " ,lastupdate = getdate()" + vbcrlf
		sql = sql & " ,lastadminid = '"& lastadminid &"' WHERE" + vbcrlf
		sql = sql & " fidx = "& fidx &""

		'response.write sql &"<br>"
		dbget.execute sql
	end if

    response.write "<script type='text/javascript'>"
    response.write "	alert('����Ǿ����ϴ�.');"
    response.write "	opener.location.reload();"
    response.write "	self.close();"
    response.write "</script>"
    dbget.close() : response.end
else
    response.write "<script type='text/javascript'>"
    response.write "	alert('�������� ��ΰ� �ƴմϴ�.');history.back();"
    response.write "</script>"
    dbget.close() : response.end
end if	
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
