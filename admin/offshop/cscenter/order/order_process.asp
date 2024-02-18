<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 고객센터
' Hieditor : 2011.03.07 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/order/order_cls.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_mail_Function_off.asp" -->
<!-- #include virtual="/lib/classes/smscls.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim orderno, mode ,buyname, buyphone, buyhp, buyemail , masteridx ,songjangdiv
dim reqname, reqphone, reqhp, reqzipcode, reqzipaddr, reqaddress, comment
dim yyyy, mm, dd, reqdate ,osms, myorderdetail,i ,isupchebeasong
dim iAsID, divcd, reguserid, title, contents_jupsu, finishuser, contents_finish
dim ipkumdiv, userid, cancelyn, emailok, smsok ,sqlStr, requiredetail, detailidx
dim songjangno ,currstate ,upcheconfirmdate ,odlvType ,itemno ,beasongdate
dim tmp , nrowCount
	beasongdate     = requestCheckVar(request("beasongdate"),30)
	itemno          = requestCheckVar(request("itemno"),10)
	odlvType        = requestCheckVar(request("odlvType"),1)
	upcheconfirmdate = requestCheckVar(request("upcheconfirmdate"),30)
	currstate       = requestCheckVar(request("currstate"),1)
	songjangno      = requestCheckVar(request("songjangno"),32)
	songjangdiv     = requestCheckVar(request("songjangdiv"),10)
	isupchebeasong  = requestCheckVar(request("isupchebeasong"),1)
	orderno = requestCheckVar(request("orderno"),16)
	mode        = requestCheckVar(request("mode"),32)
	buyname     = requestCheckVar(request("buyname"),32)
	buyphone    = requestCheckVar(request("buyphone"),16)
	buyhp       = requestCheckVar(request("buyhp"),16)
	buyemail    = requestCheckVar(request("buyemail"),128)	
	reqname     = requestCheckVar(request("reqname"),32)
	reqphone    = requestCheckVar(request("reqphone"),16)
	reqhp       = requestCheckVar(request("reqhp"),16)
	reqzipcode  = requestCheckVar(request("reqzipcode"),7)
	reqzipaddr  = requestCheckVar(request("reqzipaddr"),128)
	reqaddress  = requestCheckVar(request("reqaddress"),512)
	comment     = request("comment")
	yyyy        = requestCheckVar(request("yyyy"),4)
	mm          = requestCheckVar(request("mm"),2)
	dd          = requestCheckVar(request("dd"),2)
	reqdate     = yyyy + "-" + dd + "-" + dd
	ipkumdiv    = requestCheckVar(request("ipkumdiv"),1)
	userid      = requestCheckVar(request("userid"),32)
	emailok     = requestCheckVar(request("emailok"),1)
	smsok       = requestCheckVar(request("smsok"),1)
	requiredetail = requestCheckVar(request("requiredetail"),10)
	detailidx     = requestCheckVar(request("detailidx"),10)
	masteridx     = requestCheckVar(request("masteridx"),10)

reguserid   = session("ssbctid")
const CNEXT = " => "

//구매자 정보 수정
if (mode = "modifybuyerinfo") then
    
    dbget.beginTrans
       
    divcd   = "A900"
    title   = "주문자 정보 수정"
    
    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""
        
    sqlStr = " select top 1 IsNULL(buyname,'') as buyname"
    sqlStr = sqlStr + " ,IsNULL(buyphone,'') as buyphone"
    sqlStr = sqlStr + " ,IsNULL(buyhp,'') as buyhp"
    sqlStr = sqlStr + " ,IsNULL(buyemail,'') as buyemail"    
    sqlStr = sqlStr + " from [db_shop].dbo.tbl_shopbeasong_order_master"
    sqlStr = sqlStr + " where masteridx='" + CStr(masteridx) + "' " + VbCrlf
    
    'response.write sqlStr & "<br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        contents_jupsu = contents_jupsu & "변경 내역" & VbCrlf
        
        if (db2html(rsget("buyname"))<>buyname) then
            contents_jupsu = contents_jupsu & "주문자명: " & rsget("buyname") & CNEXT & buyname & VbCrlf
        end if
        
        if (rsget("buyphone")<>buyphone) then
            contents_jupsu = contents_jupsu & "주문자전화: " & rsget("buyphone") & CNEXT & buyphone & VbCrlf
        end if
        
        if (rsget("buyhp")<>buyhp) then
            contents_jupsu = contents_jupsu & "주문자핸드폰: " & rsget("buyhp") & CNEXT & buyhp & VbCrlf
        end if
        
        if (db2html(rsget("buyemail"))<>buyemail) then
            contents_jupsu = contents_jupsu & "주문자이메일: " & rsget("buyemail") & CNEXT & buyemail & VbCrlf
        end if
	end if
    
    rsget.Close
    
    contents_finish = contents_jupsu
    
    sqlStr = ""
    sqlStr = " update [db_shop].dbo.tbl_shopbeasong_order_master"     + VbCrlf
    sqlStr = sqlStr + " set buyname='" + html2db(buyname) + "' "   + VbCrlf
    sqlStr = sqlStr + " ,buyphone = '" + CStr(buyphone) + "' "  + VbCrlf
    sqlStr = sqlStr + " ,buyhp = '" + CStr(buyhp) + "' "        + VbCrlf
    sqlStr = sqlStr + " ,buyemail = '" + html2db(buyemail) + "' "  + VbCrlf    
    sqlStr = sqlStr + " where masteridx='" + CStr(masteridx) + "' " + VbCrlf

    'response.write sqlStr & "<br>"    
    dbget.Execute sqlStr

    ''html2db 사용하지 말것.
    iAsID = RegCSMaster_off(divcd , orderno, reguserid, title, contents_jupsu,masteridx)

    Call FinishCSMaster_off(iAsid, finishuser, html2db(contents_finish))
    
    Call AddCustomerOpenContents_off(iAsid, html2db(contents_finish))
    
    If Err.Number = 0 Then
        dbget.CommitTrans        
    Else
        dbget.RollBackTrans
        response.write "<script type='text/javascript'>alert('데이타를 저장하는 도중에 에러가 발생하였습니다')</script>"
        response.write "<script type='text/javascript'>history.back()</script>"
        dbget.close()	:	response.End
    End If
    
    Call SendCsActionMail_off(iAsID)
    
    response.write "<script type='text/javascript'>"
    response.write "	alert('수정 되었습니다.');"    
    response.write "	opener.parent.listFrame.location.reload();"
    response.write "	opener.parent.detailFrame.location.reload();"
    response.write "	window.close();"
    response.write "</script>"
    dbget.close()	:	response.End

'//배송지 정보 수정
elseif (mode="modifyreceiverinfo") then

    dbget.beginTrans

    divcd   = "A900"
    title   = "수령인 정보 수정"  
    contents_jupsu = ""
    finishuser      = reguserid
    contents_finish = ""
    
    sqlStr = " select top 1 IsNULL(reqname,'') as reqname"
    sqlStr = sqlStr + " ,IsNULL(reqphone,'') as reqphone"
    sqlStr = sqlStr + " ,IsNULL(reqhp,'') as reqhp"
    sqlStr = sqlStr + " ,IsNULL(reqzipcode,'') as reqzipcode"
    sqlStr = sqlStr + " ,IsNULL(reqzipaddr,'') as reqzipaddr"
    sqlStr = sqlStr + " ,IsNULL(reqaddress,'') as reqaddress"
    sqlStr = sqlStr + " ,IsNULL(comment,'') as comment"
    sqlStr = sqlStr + " from db_shop.dbo.tbl_shopbeasong_order_master"
    sqlStr = sqlStr + " where masteridx='" + CStr(masteridx) + "' " + VbCrlf
    
    'response.write sqlStr &"<Br>"
    rsget.Open sqlStr,dbget,1
    
    if Not rsget.Eof then
        contents_jupsu = contents_jupsu & "변경전 내역" & VbCrlf
        if (db2html(rsget("reqname"))<>reqname) then
            contents_jupsu = contents_jupsu & "수령인명: " & rsget("reqname") & CNEXT & reqname & VbCrlf
        end if
        
        if (rsget("reqphone")<>reqphone) then
            contents_jupsu = contents_jupsu & "수령인전화: " & rsget("reqphone") & CNEXT & reqphone & VbCrlf
        end if
        
        if (rsget("reqhp")<>reqhp) then
            contents_jupsu = contents_jupsu & "수령인핸드폰: " & rsget("reqhp") & CNEXT & reqhp & VbCrlf
        end if
        
        if (rsget("reqzipcode") <> reqzipcode) or (rsget("reqzipaddr") <> reqzipaddr) or (db2html(rsget("reqaddress")) <> reqaddress)  then
            contents_jupsu = contents_jupsu & "수령인주소: [" & rsget("reqzipcode") & "] " & rsget("reqzipaddr") & " " & rsget("reqaddress") & CNEXT & "[" & reqzipcode & "] " & reqzipaddr & " " & reqaddress & VbCrlf
        end if
    
        if (db2html(rsget("comment"))<>comment) then
            contents_jupsu = contents_jupsu & "기타사항: " & rsget("comment") & CNEXT & comment & VbCrlf
        end if
    end if
    
    rsget.Close
    
    contents_finish = contents_jupsu
	
	sqlStr = ""
    sqlStr = " update db_shop.dbo.tbl_shopbeasong_order_master set"     + VbCrlf
    sqlStr = sqlStr + " reqname='" + html2db(reqname) + "' "   + VbCrlf
    sqlStr = sqlStr + " ,reqphone = '" + CStr(reqphone) + "' "  + VbCrlf
    sqlStr = sqlStr + " ,reqhp = '" + CStr(reqhp) + "' "        + VbCrlf
    sqlStr = sqlStr + " ,reqzipcode = '" + CStr(reqzipcode) + "' "  + VbCrlf
    sqlStr = sqlStr + " ,reqzipaddr = '" + CStr(reqzipaddr) + "' "    + VbCrlf
    sqlStr = sqlStr + " ,reqaddress = '" + html2db(reqaddress) + "' "    + VbCrlf
    sqlStr = sqlStr + " ,comment = '" + html2db(comment) + "' "    + VbCrlf
    sqlStr = sqlStr + " where masteridx='" + CStr(masteridx) + "' " + VbCrlf
    
    'response.write sqlStr &"<Br>"
    dbget.Execute sqlStr


    ''html2db 사용하지 말것.
	iAsID = RegCSMaster_off(divcd , orderno, reguserid, title, contents_jupsu, masteridx)

	Call FinishCSMaster_off(iAsid, finishuser, html2db(contents_finish))
        
	Call AddCustomerOpenContents_off(iAsid, html2db(contents_finish))
    
    If Err.Number = 0 Then
        dbget.CommitTrans
    Else
        dbget.RollBackTrans
        response.write "<script type='text/javascript'>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n(에러코드 : " + CStr(errcode) + ")')</script>"
        response.write "<script type='text/javascript'>history.back()</script>"
        dbget.close()	:	response.End
    End If

    Call SendCsActionMail_off(iAsID)

    response.write "<script type='text/javascript'>"
    response.write "	alert('수정 되었습니다.');"    
    response.write "	opener.parent.listFrame.location.reload();"
    response.write "	opener.parent.detailFrame.location.reload();"
    response.write "	window.close();"
    response.write "</script>"
    dbget.close()	:	response.End

'//주문상품 수량 수정
elseif mode="itemno" then

	On Error resume Next
	if (upcheconfirmdate<>"") then tmp = CDate(upcheconfirmdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 확인일이 올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	if (beasongdate<>"") then tmp = CDate(beasongdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 배송일이  올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	On Error Goto 0
	    
    sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail set" + VbCrlf
	sqlStr = sqlStr + " itemno='" + CStr(itemno) + "'" + VbCrlf
	sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)  + VbCrlf
    
    'response.write sqlStr &"<br>"
	dbget.Execute sqlStr,nrowCount
       
    response.write "<script type='text/javascript'>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
	response.write "<script type='text/javascript'>location.replace('/admin/offshop/cscenter/order/orderdetailedit.asp?detailidx=" + detailidx + "');</script>"
	dbget.close()	:	response.End

'//확인상태변경
elseif mode="currstate" then

	On Error resume Next
	if (upcheconfirmdate<>"") then tmp = CDate(upcheconfirmdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 확인일이 올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	if (beasongdate<>"") then tmp = CDate(beasongdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 배송일이  올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	On Error Goto 0

    dbget.beginTrans
		
    ''/미확인
    if (currstate="") then
        sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail set" + VbCrlf
		sqlStr = sqlStr + " currstate=0"  & VbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=NULL" & VbCrlf
		sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,NULL)" & VbCrlf
		sqlStr = sqlStr + " ,songjangdiv=NULL"
		sqlStr = sqlStr + " ,songjangno=NULL"
		sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)  + VbCrlf
		
		'response.write sqlStr &"<br>"
		dbget.Execute sqlStr,nrowCount
	
	'/업체통보
	elseif (currstate="2") then
	    sqlStr = "update D set" + VbCrlf
		sqlStr = sqlStr + " D.currstate=" + CStr(currstate) + ""  & VbCrlf
		sqlStr = sqlStr + " From db_shop.dbo.tbl_shopbeasong_order_detail D" & VbCrlf
		sqlStr = sqlStr + " Join db_shop.dbo.tbl_shopbeasong_order_master M" & VbCrlf
		sqlStr = sqlStr + "     on D.masteridx=M.masteridx" & VbCrlf
		sqlStr = sqlStr + " where D.detailidx=" + CStr(detailidx)  + VbCrlf
		sqlStr = sqlStr + " and M.ipkumdiv>3"
		sqlStr = sqlStr + " and D.currstate=0"
		
		'response.write sqlStr &"<br>"
		dbget.Execute sqlStr,nrowCount
	
	'/주문확인
    elseif (currstate="3") then
        sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail set" + VbCrlf
		sqlStr = sqlStr + " currstate=" + CStr(currstate) + ""  & VbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
		sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,NULL)" & VbCrlf
		sqlStr = sqlStr + " ,songjangdiv=NULL"
		sqlStr = sqlStr + " ,songjangno=NULL"
		sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)  + VbCrlf
		
		'response.write sqlStr &"<br>"
		dbget.Execute sqlStr,nrowCount
    
    '/출고완료
    elseif (currstate="7") then
        sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail set" + VbCrlf
		sqlStr = sqlStr + " currstate=" + CStr(currstate) + ""  & VbCrlf
		sqlStr = sqlStr + " ,upcheconfirmdate=IsNULL(upcheconfirmdate,getdate())" & VbCrlf
		sqlStr = sqlStr + " ,beasongdate=IsNULL(beasongdate,getdate())" & VbCrlf
		sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)  + VbCrlf
		
		'response.write sqlStr &"<br>"
		dbget.Execute sqlStr,nrowCount
    end if

    If Err.Number = 0 Then
        dbget.CommitTrans
    	response.write "<script type='text/javascript'>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
		response.write "<script type='text/javascript'>location.replace('/admin/offshop/cscenter/order/orderdetailedit.asp?detailidx=" + detailidx + "');</script>"        
        dbget.close()	:	response.End    
    Else
        dbget.RollBackTrans
        response.write "<script type='text/javascript'>alert('데이타를 저장하는 도중에 에러가 발생하였습니다')</script>"
        response.write "<script type='text/javascript'>history.back()</script>"
        dbget.close()	:	response.End
    End If

'/택배정보수정
elseif mode="songjangdiv" then

	On Error resume Next
	if (upcheconfirmdate<>"") then tmp = CDate(upcheconfirmdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 확인일이 올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	if (beasongdate<>"") then tmp = CDate(beasongdate)
	if Err then
	    response.write "<script type='text/javascript'>alert('업체 배송일이  올바르지 않습니다.');history.back();</script>"
	    dbget.close()	:	response.End
	end if
	    
	On Error Goto 0
    
    '출고완료만 입력가능
    if currstate <> "7" then
	response.write "<script type='text/javascript'>alert('출고완료로 변경 후, 입력하세요.');</script>"
	response.write "<script type='text/javascript'>history.back();</script>"
	dbget.close()	:	response.End
    end if
 	    
    '택배사수정
    sqlStr = "update db_shop.dbo.tbl_shopbeasong_order_detail" + VbCrlf
	sqlStr = sqlStr + " set songjangdiv='" + CStr(songjangdiv) + "'" + VbCrlf
	sqlStr = sqlStr + " ,songjangno='" + CStr(songjangno) + "'" + VbCrlf
	sqlStr = sqlStr + " where detailidx=" + CStr(detailidx)  + VbCrlf

	'response.write sqlStr &"<br>"
	dbget.Execute sqlStr,nrowCount

	response.write "<script type='text/javascript'>alert('" + CStr(nrowCount) + "개 수정 되었습니다.');</script>"
	response.write "<script type='text/javascript'>location.replace('/admin/offshop/cscenter/order/orderdetailedit.asp?detailidx=" + detailidx + "');</script>"
	dbget.close()	:	response.End
 		
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
