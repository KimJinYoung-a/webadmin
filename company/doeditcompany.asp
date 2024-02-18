<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->

<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim userid
dim txcompanyname, txpassword
dim txaddress1, txaddress2
dim txurl, txmanagername
dim txphone, txfax, txemail
dim txnewpassword1, txnewpassword2

userid = session("ssBctId")
txcompanyname = html2db(request("txcompanyname"))
txpassword = request("txpassword")
txaddress1 = html2db(request("txaddress1"))
txaddress2 = html2db(request("txaddress2"))
txurl = html2db(request("txurl"))
txmanagername = html2db(request("txmanagername"))
txphone = request("txphone")
txfax = request("txfax")
txemail = request("txemail")
txnewpassword1 = request("txnewpassword1")
txnewpassword2 = request("txnewpassword2")

''if txnewpassword1<>"" then txpassword = txnewpassword1

dim sqlstr
dim orgpass
sqlstr = "select top 1 Enc_password, Enc_password64 from [db_partner].[dbo].tbl_partner"
sqlstr = sqlstr + " where id='" + userid + "'"
rsget.Open sqlStr,dbget,1
    orgpass = rsget("Enc_password64")  ''''orgpass = rsget("Enc_password")
rsget.Close


if (UCASE(orgpass)<>UCASE(SHA256(md5(txpassword)))) then ''''if (UCASE(orgpass)<>UCASE((md5(txpassword)))) then
%>
<script language='javascript'>
alert('비밀번호가 일치하지 않습니다.');
location.replace('<%= refer %>');
</script>
<%
dbget.close()	:	response.End
end if

sqlstr = "update [db_partner].[dbo].tbl_partner" + VbCrlf
sqlstr = sqlstr + " set lastInfoChgDT=getdate(), Enc_password64='" + SHA256(md5(txnewpassword1)) + "'" + VbCrlf
sqlstr = sqlstr + " ,Enc_password=''" + VbCrlf
sqlstr = sqlstr + " ,company_name='" + txcompanyname + "'" + VbCrlf
sqlstr = sqlstr + " ,address='" + txaddress1 + "'" + VbCrlf
sqlstr = sqlstr + " ,tel='" + txphone + "'" + VbCrlf
sqlstr = sqlstr + " ,fax='" + txfax +"'" + VbCrlf
sqlstr = sqlstr + " ,url='" + txurl +"'" + VbCrlf
sqlstr = sqlstr + " ,manager_name='" + txmanagername + "'" + VbCrlf
sqlstr = sqlstr + " ,manager_address='" + txaddress2 + "'" + VbCrlf
sqlstr = sqlstr + " ,email='" + txemail + "'" + VbCrlf
sqlstr = sqlstr + " where id='" + userid + "'"

'response.write sqlStr
rsget.Open sqlStr,dbget,1


''최종 로그인 일자 저장 //2014/07/14 '' tbl_user_tenbyten 사번로그인 제외
sqlStr = "exec db_partner.dbo.sp_Ten_Add_PartnerLoginLog '"&userid&"','"&Left(request.ServerVariables("REMOTE_ADDR"),16)&"','R','',0"
dbget.Execute sqlStr
    
%>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->