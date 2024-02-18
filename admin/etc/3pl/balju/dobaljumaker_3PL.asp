<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 80
%>
<%
'###########################################################
' Description : 출고지시
' History : 이상구 생성
'			2018.03.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
function FormatStr(n,orgData)
	dim tmp
	if (n-Len(CStr(orgData))) < 0 then
		FormatStr = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	FormatStr = tmp
end Function

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim orderserial
dim differencekey, workgroup, songjangdiv
dim tplcompanyid

mode        = request("mode")
orderserial = request("orderserial")
workgroup   = request("workgroup")
tplcompanyid   = request("extSiteName")
songjangdiv = request("songjangdiv")

dummiserial = orderserial
orderserial = split(orderserial,"|")

dim sqlStr,i
dim iid
dim dummiserial
dim obaljudate
dim errcode

if mode="arr" then
	''유효성체크.
	dummiserial = Mid(dummiserial,2,Len(dummiserial))
	dummiserial = replace(dummiserial,"|","','")
	sqlStr = " select top 1 orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " where orderserial in ('" + dummiserial + "')"

	rsget_TPL.Open sqlStr,dbget_TPL,1
	if Not rsget_TPL.Eof then
		response.write "<script language='javascript'>"
		response.write "alert('" + rsget_TPL("orderserial") + " : 이미 출고지시된 주문입니다. \n\n이미 출고지시한 주문건은 중복 출고지시 할 수 없습니다.');"
		response.write "location.replace('" + CStr(refer) + "');"
		response.write "</script>"
		dbget_TPL.close()	:	response.End
	end if
	rsget_TPL.Close

	sqlStr = " select (count(idx) + 1) as differencekey"
	sqlStr = sqlStr + " from [db_threepl].[dbo].[tbl_tpl_balju_master]"
	sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21)"

	rsget_TPL.Open sqlStr,dbget_TPL,1
		differencekey = rsget_TPL("differencekey")
	rsget_TPL.close


On Error Resume Next
dbget_TPL.beginTrans

If Err.Number = 0 Then
        errcode = "001"


    '' 택배사별 출고지시 system으로 수정

	'#######출고지시 마스터###############
	sqlStr = "insert into [db_threepl].[dbo].[tbl_tpl_balju_master](baljudate,differencekey,workgroup, songjangdiv, tplcompanyid)"
	sqlStr = sqlStr + " values(getdate()," + CStr(differencekey) + ",'" + workgroup + "','" + songjangdiv + "', '" + CStr(tplcompanyid) + "')"
	rsget_TPL.Open sqlStr,dbget_TPL,1

	sqlStr = "select top 1 idx, convert(varchar(19),baljudate,21) as baljudate from [db_threepl].[dbo].[tbl_tpl_balju_master] order by idx desc"
	rsget_TPL.Open sqlStr,dbget_TPL,1
	    iid = rsget_TPL("idx")
	    obaljudate = rsget_TPL("baljudate")
	rsget_TPL.Close

	'#######출고지시 디테일###############
	for i=0 to Ubound(orderserial)
		if orderserial(i)<>"" then
			sqlStr = "insert into [db_threepl].[dbo].[tbl_tpl_balju_detail](baljuid,orderserial)"
			sqlStr = sqlStr + " values(" + CStr(iid) + ","
			sqlStr = sqlStr + " '" + orderserial(i) + "'"
			sqlStr = sqlStr + " )"
			rsget_TPL.Open sqlStr,dbget_TPL,1
		end if
	next

    ''** [db_threepl].[dbo].[tbl_tpl_balju_detail].baljusongjangno is NULL 인경우 업체배송으로 인식 (Logics 기존 시스템)
    ''텐바이텐 배송의 경우 송장번호를 not null 값으로 입력..

    sqlStr = "update [db_threepl].[dbo].[tbl_tpl_balju_detail]" + VbCrlf
	sqlStr = sqlStr + " set baljusongjangno=''"
	sqlStr = sqlStr + " where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " and baljusongjangno is NULL "
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "002"

	''' 주문내역 출고지시일자 입력
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]"
	sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where ipkumdiv > 4 "
	sqlStr = sqlStr + " and baljudate is NULL"
	sqlStr = sqlStr + " and orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	rsget_TPL.Open sqlStr,dbget_TPL,1


	'#######주문 통보 진행 ############### (출고지시일 지정)
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderMaster]"
	sqlStr = sqlStr + " set ipkumdiv='5'"
	sqlStr = sqlStr + " ,baljudate='" + CStr(obaljudate) + "'"
	sqlStr = sqlStr + " where ipkumdiv=4"
	sqlStr = sqlStr + " and orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "004"
	'###### Order Detail 텐바이텐 배송 출고지시일 저장 ############
	sqlStr = "update [db_threepl].[dbo].[tbl_tpl_orderDetail]"
	sqlStr = sqlStr + " set currstate='2'"
	sqlStr = sqlStr + " ,songjangdiv=" + CStr(songjangdiv)
	sqlStr = sqlStr + " where orderserial in "
	sqlStr = sqlStr + " (select orderserial from [db_threepl].[dbo].[tbl_tpl_balju_detail]"
	sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " and [db_threepl].[dbo].[tbl_tpl_orderDetail].itemid<>0"
	rsget_TPL.Open sqlStr,dbget_TPL,1
end if


If Err.Number = 0 Then
        errcode = "013"
    '' 재고 업데이트
    ''sqlStr = " exec [db_summary].[dbo].sp_Ten_RealtimeStock_balju " & iid
    ''dbget_TPL.execute sqlStr
end if


If Err.Number = 0 Then
    dbget_TPL.CommitTrans
Else
    dbget_TPL.RollBackTrans
    response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n관리자 문의 요망 (에러코드 : " + CStr(errcode) + ")');</script>"
    response.write "<script>history.back()</script>"
    dbget_TPL.close()	:	dbget.close()	:	response.End
End If
on error Goto 0

end if

%>

<script language="javascript">
alert('출고지시서가 생성 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLclose.asp" -->