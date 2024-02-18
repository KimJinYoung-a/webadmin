<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  현장수령주문
' History : 2012.05.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->

<%
dim orderserial , mode ,ojumun ,isFinishValid , sqlstr, userid
	orderserial = requestCheckVar(request("orderserial"),11)
	mode = requestCheckVar(request("mode"),32)

isFinishValid= false

'/현장수령주문완료처리	
if mode = "siteReceivefinsh" then
	if orderserial = "" then
		response.write "<script language='javascript'>"
		response.write "	alert('주문번호가 없습니다');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	'/주문조회
	set ojumun = new CJumunMaster
		ojumun.FRectOrderSerial = orderserial
		ojumun.SearchJumunList

	if ojumun.ftotalcount < 1 then
		response.write "<script language='javascript'>"
		response.write "	alert('주문건이 없습니다');"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	
	'/상태값이 결제완료고 , 완료처리 이전이고
	isFinishValid = (ojumun.FMasterItemList(0).FIpkumdiv>3) and (ojumun.FMasterItemList(0).FIpkumdiv<8)
	
	'/최소일경우
	isFinishValid = isFinishValid and (ojumun.FMasterItemList(0).FCancelyn="N")
	
    '/cs 메모에 저장하기위해 
    userid = ojumun.FMasterItemList(0).FUserID
    
	if (Not isFinishValid) then
		if (ojumun.FMasterItemList(0).FIpkumdiv>7) then
			response.write "<script language='javascript'>"
			response.write "	alert('이미 처리 완료된 주문 입니다.');"
			response.write "</script>"
			dbget.close()	:	response.End
		elseif (ojumun.FMasterItemList(0).FIpkumdiv<4) then
			response.write "<script language='javascript'>"
			response.write "	alert('결제이전 주문건 입니다.');"
			response.write "</script>"
			dbget.close()	:	response.End
		elseif (ojumun.FMasterItemList(0).FCancelyn<>"N") then
			response.write "<script language='javascript'>"
			response.write "	alert('취소된 주문 입니다.');"
			response.write "</script>"
			dbget.close()	:	response.End			
		elseif (ojumun.FMasterItemList(0).Fjumundiv<>"7") then
			response.write "<script language='javascript'>"
			response.write "	alert('현장수령 주문건이 아닙니다.');"
			response.write "</script>"
			dbget.close()	:	response.End
		end if
	end if
	
    
    
	set ojumun = nothing

	sqlStr = "update D" + vbcrlf
	sqlStr = sqlStr + " set currstate='7'" + vbcrlf
	sqlStr = sqlStr + " ,songjangno='기타출고'" + vbcrlf
	sqlStr = sqlStr + " ,songjangdiv='99'" + vbcrlf
	sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" + vbcrlf
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail D" + vbcrlf
	sqlStr = sqlStr + " Join [db_order].[dbo].tbl_order_master m" + vbcrlf
    sqlStr = sqlStr + " 	on m.orderserial=d.orderserial" + vbcrlf
	sqlStr = sqlStr + " where d.orderserial = '"&orderserial&"'" + vbcrlf
	sqlStr = sqlStr + " and d.cancelyn<>'Y'" + vbcrlf
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and m.jumundiv='7'"
	
	'response.write sqlstr & "<Br>"
	dbget.execute sqlstr

    sqlStr = "update [db_order].[dbo].tbl_order_master" + vbcrlf
    sqlStr = sqlStr + " set ipkumdiv='8'" + vbcrlf
    sqlStr = sqlStr + " , beadaldate=getdate()" + vbcrlf
	sqlstr = sqlstr & " ,songjangdiv = '99'" + vbcrlf
	sqlstr = sqlstr & " ,deliverno = '기타출고'" + vbcrlf
    sqlStr = sqlStr + " where orderserial in (" + vbcrlf
    sqlStr = sqlStr + "     select m.orderserial" + vbcrlf
    sqlStr = sqlStr + "     from [db_order].[dbo].tbl_order_master m" + vbcrlf
    sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_order_detail d" + vbcrlf
    sqlStr = sqlStr + "         on m.orderserial=d.orderserial" + vbcrlf
    sqlStr = sqlStr + "     where m.orderserial = '"&orderserial&"'" + vbcrlf
    sqlStr = sqlStr + "     and m.cancelyn='N'" + vbcrlf
    sqlStr = sqlStr + "     and m.jumundiv<>9" + vbcrlf
    sqlStr = sqlStr + "     and d.itemid<>0" + vbcrlf
    sqlStr = sqlStr + "     and d.cancelyn<>'Y'" + vbcrlf
	sqlStr = sqlStr + " 	and m.jumundiv='7'"    
    sqlStr = sqlStr + "     group by m.orderserial" + vbcrlf
    sqlStr = sqlStr + "     having sum(case when IsNull(d.currstate,'0')<>'7' then 1 else 0 end )=0"
    sqlStr = sqlStr + " ) "
	
	'response.write sqlstr & "<Br>"
	dbget.execute sqlstr

    ''CS 메모 저장 // 서동석 추가
    sqlStr = " insert into [db_cs].[dbo].tbl_cs_memo(orderserial, divcd, userid, mmgubun, qadiv, phoneNumber, writeuser, finishuser, contents_jupsu, finishyn,finishdate,regdate) "
    sqlStr = sqlStr + " values('" + CStr(orderserial) + "','1','" + CStr(userid) + "','0','20','','" + session("ssBctId") + "','" + session("ssBctId") + "','현장수령 완료진행','Y',getdate(),getdate()) "
    dbget.Execute sqlStr
        
	response.write "<script language='javascript'>"
	''response.write "	parent.opener.location.reload();"	
	response.write "	alert('출고완료 처리 되었습니다.');"
	response.write "	parent.location.href='/admin/apps/siteReceive/popSiteReceive.asp?orderserial="&orderserial&"&aplot=Y';"
	''response.write "	parent.plotReceipt();"
	response.write "</script>"
	dbget.close()	:	response.End

else
	response.write "<script language='javascript'>"
	response.write "	alert('구분자가 없습니다');"
	response.write "</script>"
	dbget.close()	:	response.End
end if	
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->