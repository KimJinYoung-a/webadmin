<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 송장입력
' History : 최초생성자모름
'			2017.04.10 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim makerid, paramgubunname, chargeArrival
makerid = requestCheckVar(request("makerid"),32)
paramgubunname = requestCheckVar(request("paramgubunname"),32)
chargeArrival = requestCheckVar(request("chargeArrival"),1)

if (paramgubunname="") then paramgubunname="반품"

if (makerid="") then
    response.write "makerid가 지정되지 않았습니다."
    dbget.close()	:	response.End
end if

dim sqlStr, AssignedCount

sqlStr = "insert into [db_sitemaster].[dbo].tbl_etc_songjang"
sqlStr = sqlStr + "( gubuncd    " + VbCrlf
sqlStr = sqlStr + ", gubunname  " + VbCrlf
sqlStr = sqlStr + ", userid     " + VbCrlf
sqlStr = sqlStr + ", username   " + VbCrlf
sqlStr = sqlStr + ", reqname    " + VbCrlf
sqlStr = sqlStr + ", reqphone   " + VbCrlf
sqlStr = sqlStr + ", reqhp      " + VbCrlf
sqlStr = sqlStr + ", reqzipcode " + VbCrlf
sqlStr = sqlStr + ", reqaddress1" + VbCrlf
sqlStr = sqlStr + ", reqaddress2" + VbCrlf
sqlStr = sqlStr + ", reqetc     " + VbCrlf
sqlStr = sqlStr + ", inputdate  " + VbCrlf
sqlStr = sqlStr + ", prizetitle " + VbCrlf
sqlStr = sqlStr + ", isupchebeasong " + VbCrlf
sqlStr = sqlStr + ", reqdeliverdate " + VbCrlf
sqlStr = sqlStr + ", chargeArrival " + VbCrlf
sqlStr = sqlStr + ") " + VbCrlf
sqlStr = sqlStr + " select top 1 " + VbCrlf
sqlStr = sqlStr + " '90'" + VbCrlf
sqlStr = sqlStr + " ,'" + paramgubunname + "'" + VbCrlf
sqlStr = sqlStr + " ,''" + VbCrlf
sqlStr = sqlStr + " ,convert(varchar(32),c.socname_kor)" + VbCrlf
sqlStr = sqlStr + " ,convert(varchar(32),c.socname_kor + '(' + p.deliver_name + ')')" + VbCrlf
sqlStr = sqlStr + " ,p.deliver_phone" + VbCrlf
sqlStr = sqlStr + " ,p.deliver_hp" + VbCrlf
sqlStr = sqlStr + " ,p.return_zipcode" + VbCrlf
sqlStr = sqlStr + " ,p.return_address" + VbCrlf
sqlStr = sqlStr + " ,p.return_address2" + VbCrlf
sqlStr = sqlStr + " ,c.socname_kor + ' " + paramgubunname + "'" + VbCrlf
sqlStr = sqlStr + " ,getdate() " + VbCrlf
sqlStr = sqlStr + " ,c.socname_kor + ' " + paramgubunname + "'" + VbCrlf
sqlStr = sqlStr + " ,'N'" + VbCrlf
sqlStr = sqlStr + " ,convert(varchar(10),getdate(),21)" + VbCrlf
sqlStr = sqlStr + " ,'" & chargeArrival & "'" + VbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p" + VbCrlf
sqlStr = sqlStr + " ,[db_user].[dbo].tbl_user_c c" + VbCrlf
sqlStr = sqlStr + " where p.id='" + makerid + "'" + VbCrlf
sqlStr = sqlStr + " and p.id=c.userid" + VbCrlf
sqlStr = sqlStr + " and p.return_zipcode is Not NULL" + VbCrlf
sqlStr = sqlStr + " and p.return_zipcode <>''" + VbCrlf

dbget.Execute sqlStr, AssignedCount

if AssignedCount>0 then
    response.write "<script type='text/javascript'>alert('등록되었습니다. - " + makerid + "');</script>"
    response.write "<script type='text/javascript'>window.close();</script>"
else
    response.write "<script type='text/javascript'>alert('등록오류. - " + makerid + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
