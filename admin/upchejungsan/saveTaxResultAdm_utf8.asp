<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 세금계산서 발행후 저장
' History : 서동석 생성
'           2022.11.08 한용민 수정(빌36524 플래시연동발행 에서 위하고 api 연동으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/classes/jungsan/electaxcls.asp" -->
<%
dim sqlStr, IsBill36524, IsWEHAGO
dim idx : idx = request("idx")
dim result : result = request("result")
dim no_tax : no_tax = request("no_tax")
dim result_msg : result_msg = request("result_msg")
dim jungsangubun : jungsangubun = request("jungsangubun")
dim write_date   : write_date = request("write_date")
dim jungsanid    : jungsanid = request("jungsanid")
dim no_iss       : no_iss = request("no_iss")
dim billsiteCode : billsiteCode = request("billsiteCode")
dim isauto : isauto = request("isauto")
IsBill36524 = (billsiteCode="B")
IsWEHAGO = (billsiteCode="WE")
no_iss = Replace(no_iss,"-","")

''기존타입과의 호환성.
if (result="00000") then
    result_msg = "OK"
else
    result_msg = "["&result&"]"&result_msg
end if

if (no_tax="") then no_tax="null"
sqlStr = " update [db_jungsan].[dbo].tbl_tax_history_master" + vbCrlf
sqlStr = sqlStr + " set tax_no='" + no_tax + "'" + vbCrlf
sqlStr = sqlStr + " , resultmsg=convert(varchar(128),'" + result_msg + "')" + vbCrlf
sqlStr = sqlStr + " where idx=" + CStr(idx) + vbCrlf

dbget.Execute sqlStr

if (result_msg="OK") then
	if (jungsangubun="ON") or (jungsangubun="AC") then
		sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master" + vbCrlf
		sqlStr = sqlStr + " set taxlinkidx=" + CStr(idx) + vbCrlf
		sqlStr = sqlStr + " ,neotaxno='" + CStr(no_tax) + "'" + vbCrlf
		sqlStr = sqlStr + " ,eseroEvalSeq='" + CStr(no_iss) + "'" + vbCrlf
		sqlStr = sqlStr + " ,billsiteCode='" + CStr(billsiteCode) + "'" + vbCrlf
		sqlStr = sqlStr + " ,finishflag='3'"  + vbCrlf
		sqlStr = sqlStr + " ,taxinputdate=getdate()"  + vbCrlf
		sqlStr = sqlStr + " ,taxregdate='" + write_date + "'"  + vbCrlf
		sqlStr = sqlStr + " where id=" + CStr(jungsanid)
''rw sqlStr
		rsget.Open sqlStr,dbget,1
	elseif (jungsangubun="OFF") or (jungsangubun="FRN") then
	''사용안함.
		sqlStr = " update [db_shop].[dbo].tbl_shop_jungsanmaster" + vbCrlf
		sqlStr = sqlStr + " set taxlinkidx=" + CStr(idx) + vbCrlf
		sqlStr = sqlStr + " ,neotaxno='" + CStr(no_tax) + "'" + vbCrlf
		sqlStr = sqlStr + " ,eseroEvalSeq='" + CStr(no_iss) + "'" + vbCrlf
		sqlStr = sqlStr + " ,billsiteCode='" + CStr(billsiteCode) + "'" + vbCrlf
		sqlStr = sqlStr + " ,currstate='3'"  + vbCrlf
		sqlStr = sqlStr + " ,taxregdate=getdate()"  + vbCrlf
		sqlStr = sqlStr + " ,segumil='" + write_date + "'"  + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(jungsanid)
''rw sqlStr
		rsget.Open sqlStr,dbget,1
    elseif (jungsangubun="OF") then
		sqlStr = " update [db_jungsan].[dbo].tbl_off_jungsan_master" + vbCrlf
		sqlStr = sqlStr + " set taxlinkidx=" + CStr(idx) + vbCrlf
		sqlStr = sqlStr + " ,neotaxno='" + CStr(no_tax) + "'" + vbCrlf
		sqlStr = sqlStr + " ,eseroEvalSeq='" + CStr(no_iss) + "'" + vbCrlf
		sqlStr = sqlStr + " ,billsiteCode='" + CStr(billsiteCode) + "'" + vbCrlf
		sqlStr = sqlStr + " ,finishflag='3'"  + vbCrlf
		sqlStr = sqlStr + " ,taxinputdate=getdate()"  + vbCrlf
		sqlStr = sqlStr + " ,taxregdate='" + write_date + "'"  + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(jungsanid)

		rsget.Open sqlStr,dbget,1

    elseif (jungsangubun="OFFSHOP") then	' 오프 가맹점
		sqlStr = " update db_shop.dbo.tbl_fran_meachuljungsan_master" + vbCrlf
		sqlStr = sqlStr + " set taxlinkidx=" + CStr(idx) + vbCrlf
		sqlStr = sqlStr + " ,neotaxno='" + CStr(no_tax) + "'" + vbCrlf
		sqlStr = sqlStr + " ,stateCd='4'"  + vbCrlf		' 계산서발행완료
		sqlStr = sqlStr + " ,taxdate = '" + write_date + "'"  + vbCrlf
		sqlStr = sqlStr + " ,taxregdate = getdate() " + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(jungsanid)

		dbget.Execute(sqlStr)

		sqlStr = " UPDATE A " + vbCrlf
		sqlStr = sqlStr + " SET a.segumDate = '" + write_date + "'"  + vbCrlf
		sqlStr = sqlStr + " FROM db_storage.dbo.tbl_ordersheet_master a " & vbCrLf
		sqlStr = sqlStr + " INNER JOIN [db_shop].[dbo].tbl_fran_meachuljungsan_submaster b " & vbCrLf
		sqlStr = sqlStr + " ON a.baljucode = b.code02 " & vbCrLf
		sqlStr = sqlStr + " WHERE b.masterIdx = " + CStr(jungsanid)

		dbget.Execute(sqlStr)

    end if
    if IsWEHAGO then
		if (isauto<>"") then
			session.codePage = 949
			response.write "<script type='text/javascript'>parent.reActEval();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script type='text/javascript'>"
			response.write "    alert('계산서가 발행 되었습니다.');"
			session.codePage = 949
			response.write "    parent.window.close();"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
    else
		if (isauto<>"") then
			session.codePage = 949
			response.write "<script type='text/javascript'>parent.reActEval();</script>"
			dbget.close()	:	response.End
		else
			response.write "<script type='text/javascript'>alert('계산서가 발행 되었습니다. ');</script>"
			session.codePage = 949
			response.write "<script type='text/javascript'>parent.closeMe();</script>"
			dbget.close()	:	response.End
		end if
    end if
else
    ''본창에서 alert
    ''response.write "<script>alert('" & Replace(result_msg,VbCrlf,"\n") & "');</script>"
end if

session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->