<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ���� ���� ���Ѽ���
' Hieditor : 2011.01.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->

<%
dim empno ,shopid , mode , sql ,tmpvalue , tmpcnt
	empno = request("empno")
	shopid = request("shopid")
	mode = request("mode")	

dim ref
	ref = request.ServerVariables("HTTP_REFERER")

tmpcnt = 0
	
'//��ǥ�����庯��
if mode = "shopfirstchange" then

	if empno = "" or shopid = "" then
		response.write "<script language='javascript'>"
		response.write " 	alert('�����ȣ�� shopid�� �����ϴ�');"
		response.write "	self.close();"
		response.write "</script>"
		response.end
	end if
	
	dbget.beginTrans

	sql = "update db_partner.dbo.tbl_partner_shopuser set" + vbcrlf
	sql = sql & " firstisusing='N'" + vbcrlf
	sql = sql & " where empno ='"&empno&"'" + vbcrlf
	
	'response.write sql &"<Br>"
	dbget.execute sql
	
	sql = ""
	sql = "update db_partner.dbo.tbl_partner_shopuser set" + vbcrlf
	sql = sql & " firstisusing='Y'" + vbcrlf
	sql = sql & " where empno ='"&empno&"' and shopid='"&shopid&"'" + vbcrlf
	
	'response.write sql &"<Br>"
	dbget.execute sql	
	
	If Err.Number = 0 Then
	    dbget.CommitTrans
	Else
	    dbget.RollBackTrans
	End If	
	
	response.write "<script language='javascript'>"
	response.write " 	alert('OK');"
	response.write " 	opener.location.reload();"
	response.write "	location.href='/common/offshop/member/shopuser_reg.asp?empno="&empno&"';"
	response.write "</script>"
	response.end

'//�����߰�
elseif mode = "shopmemberadd" then
	
	sql = "select count(*) as cnt"
	sql = sql & " from db_partner.dbo.tbl_partner_shopuser"
	sql = sql & " where shopid = '"&shopid&"' and empno = '"&empno&"'"
	
	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	
	if not rsget.EOF  then        
		tmpvalue = rsget("cnt") > 0
	end if
	
	rsget.close
	
	if tmpvalue then		
		response.write "<script language='javascript'>"
		response.write " 	alert('�ش���忡 ���� ������ �̹� ��ϵǾ� �ֽ��ϴ�.');"
		response.write "	location.href='/common/offshop/member/shopuser_reg.asp?empno="&empno&"';"
		response.write "</script>"
		response.end
	else
		sql = ""
		sql = "insert into db_partner.dbo.tbl_partner_shopuser (empno ,shopid ,firstisusing)" + vbcrlf
		sql = sql & " 	select ut.empno ,'"&shopid&"'" + vbcrlf
		sql = sql & "	,(case when isnull(t.cnt,0) > 0 then 'N' else 'Y' end)" + vbcrlf
		sql = sql & "	from db_partner.dbo.tbl_user_tenbyten ut" + vbcrlf
		sql = sql & "	left join (" + vbcrlf
		sql = sql & "		select empno , count(*) as cnt" + vbcrlf
		sql = sql & "		from db_partner.dbo.tbl_partner_shopuser" + vbcrlf	
		sql = sql & "		where empno = '"&empno&"'" + vbcrlf
		sql = sql & "		group by empno" + vbcrlf
		sql = sql & "	) as t" + vbcrlf
		sql = sql & "	on ut.empno = t.empno" + vbcrlf
		sql = sql & "	where ut.empno = '"&empno&"'" + vbcrlf
		
		'response.write sql &"<Br>"
		dbget.execute sql
		
		response.write "<script language='javascript'>"
		response.write " 	alert('OK');"
		response.write " 	opener.location.reload();"
		response.write "	location.href='/common/offshop/member/shopuser_reg.asp?empno="&empno&"';"
		response.write "</script>"
		response.end
		
	end if

'//����
elseif mode = "del" then

	sql = "select su.firstisusing"
	sql = sql & " ,(select count(*) as cnt from db_partner.dbo.tbl_partner_shopuser where su.empno = empno) as cnt"
	sql = sql & " from db_partner.dbo.tbl_partner_shopuser su"
	sql = sql & " where su.shopid = '"&shopid&"' and su.empno = '"&empno&"'"
	
	'response.write sql &"<Br>"
	rsget.Open sql,dbget,1
	
	if not rsget.EOF  then        
		tmpvalue = rsget("firstisusing")
		tmpcnt = rsget("cnt")
	end if
	
	rsget.close

	if tmpvalue = "Y" and tmpcnt > 2 then
		response.write "<script language='javascript'>"
		response.write " 	alert('��ǥ���� ������, ���������� 2�� �̻� �� �����մϴ�.\n��ǥ�������� �ٸ� �������� ������ ���� �ϼ���.');"
		response.write "	location.href='/common/offshop/member/shopuser_reg.asp?empno="&empno&"';"
		response.write "</script>"
		response.end
	else
		sql = "delete from db_partner.dbo.tbl_partner_shopuser" + vbcrlf
		sql = sql & " where shopid = '"&shopid&"' and empno = '"&empno&"'"

		'response.write sql &"<Br>"
		dbget.execute sql

		sql = "select su.firstisusing"
		sql = sql & " ,(select count(*) as cnt from db_partner.dbo.tbl_partner_shopuser where su.empno = empno) as cnt"
		sql = sql & " from db_partner.dbo.tbl_partner_shopuser su"
		sql = sql & " where su.empno = '"&empno&"'"
		
		'response.write sql &"<Br>"
		rsget.Open sql,dbget,1
		
		if not rsget.EOF  then        
			tmpvalue = rsget("firstisusing")
			tmpcnt = rsget("cnt")
		end if
		
		rsget.close
		
		'//��ǥ���� ������ ���������� 1�� ������� ��ǥ�������� ����ħ
		sql = " update u set" + vbcrlf
		sql = sql & " u.firstisusing='Y'" + vbcrlf
		sql = sql & " from db_partner.dbo.tbl_partner_shopuser u" + vbcrlf
		sql = sql & " left join (" + vbcrlf
		sql = sql & " 	select" + vbcrlf
		sql = sql & " 	su.empno, count(*) as cnt" + vbcrlf
		sql = sql & " 	from db_partner.dbo.tbl_partner_shopuser su" + vbcrlf
		sql = sql & " 	left join (" + vbcrlf
		sql = sql & " 		select" + vbcrlf
		sql = sql & " 		empno" + vbcrlf
		sql = sql & " 		from db_partner.dbo.tbl_partner_shopuser" + vbcrlf
		sql = sql & " 		where firstisusing='Y' and empno = '"&empno&"'" + vbcrlf
		sql = sql & " 		group by empno" + vbcrlf
		sql = sql & " 	) suu" + vbcrlf
		sql = sql & " 		on su.empno = suu.empno" + vbcrlf
		sql = sql & " 	where suu.empno is null and su.empno = '"&empno&"'" + vbcrlf
		sql = sql & " 	group by su.empno" + vbcrlf
		sql = sql & " 	having count(*) = 1" + vbcrlf
		sql = sql & " ) as t" + vbcrlf
		sql = sql & " 	on u.empno = t.empno" + vbcrlf
		sql = sql & " where t.empno is not null and u.empno = '"&empno&"'"
		
		'response.write sql &"<br>"
		dbget.execute sql
		
		response.write "<script language='javascript'>"
		response.write " 	alert('OK');"
		response.write " 	opener.location.reload();"
		response.write "	location.href='/common/offshop/member/shopuser_reg.asp?empno="&empno&"';"
		response.write "</script>"
		response.end
			
	end if

end if
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->