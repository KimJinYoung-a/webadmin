<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  사은품 처리
' History : 2010.03.11 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim mode, strSql,strSqlAdd , giftkind_code
dim evt_code, gift_name , gift_startdate , gift_enddate , gift_status , gift_type , gift_range1
dim gift_range2 , giftkind_cnt , giftkind_type , giftkind_limit , opendate , closedate
dim gift_scope , gift_code , gift_using , strAdd , giftkind_givecnt
dim gift_scope_add, giftkind_limit_sold
dim gift_itemname

dim itemgubun , itemoption , shopitemid
dim giftkind_name, gift_itemgubun , gift_itemoption , gift_shopitemid
dim makerid

	mode = requestCheckVar(Request.Form("mode"),25)
	giftkind_name 	= html2db(requestCheckVar(Request.Form("giftkind_name"),64))

	shopitemid	= requestCheckVar(Request.Form("shopitemid"),10)
	itemgubun = requestCheckVar(Request.Form("itemgubun"),2)
	itemoption = requestCheckVar(Request.Form("itemoption"),4)

	gift_shopitemid	= requestCheckVar(Request.Form("gift_shopitemid"),10)
	gift_itemgubun = requestCheckVar(Request.Form("gift_itemgubun"),2)
	gift_itemoption = requestCheckVar(Request.Form("gift_itemoption"),4)

	giftkind_code = requestCheckVar(Request.Form("giftkind_code"),10)
	evt_code			= requestCheckVar(Request.Form("evt_code"),10)
	gift_name = html2db(requestCheckVar(Request.Form("gift_name"),64))
	gift_startdate = requestCheckVar(Request.Form("gift_startdate"),10)
	gift_enddate = requestCheckVar(Request.Form("gift_enddate"),10)
	gift_status = requestCheckVar(Request.Form("gift_status"),4)
	gift_type = requestCheckVar(Request.Form("gift_type"),10)
	gift_range1	= requestCheckVar(Request.Form("gift_range1"),10)
	gift_range2	= requestCheckVar(Request.Form("gift_range2"),10)
	giftkind_cnt = requestCheckVar(Request.Form("giftkind_cnt"),10)
	giftkind_type = requestCheckVar(Request.Form("giftkind_type"),10)
	giftkind_limit = requestCheckVar(Request.Form("giftkind_limit"),10)
	gift_scope = requestCheckVar(Request.Form("gift_scope"),10)
	opendate = requestCheckVar(Request.Form("opendate"),30)
	closedate = requestCheckVar(Request.Form("closedate"),30)
	gift_code = requestCheckVar(Request.Form("gift_code"),10)
	gift_using = requestCheckVar(Request.Form("gift_using"),10)
	giftkind_givecnt	= requestCheckVar(Request.Form("giftkind_givecnt"),10)		'/사은품수량
	menupos = requestCheckVar(request("menupos"),10)

	gift_scope_add = html2db(requestCheckVar(Request.Form("gift_scope_add"),256))
	giftkind_limit_sold = requestCheckVar(Request.Form("giftkind_limit_sold"),30)

	makerid = requestCheckVar(Request.Form("makerid"),32)

	gift_itemname = requestCheckVar(Request.Form("gift_itemname"),40)

	'response.write mode &"<br>"

'//사은품 상품 등록
if mode = "giftitemedit" then

	'사용안함(skyer9)

'	'//수정
'	if giftkind_code <> "" then
'
'		IF itemid <> "" THEN
'		strSql = "SELECT shopitemid FROM [db_shop].dbo.tbl_shop_item where" + vbcrlf
'		strSql = strSql & " shopitemid = "&itemid&" and itemgubun = '"&itemgubun&"'" + vbcrlf
'		strSql = strSql & " and itemoption = '"&itemoption&"'" + vbcrlf
'
'		rsget.Open strSql, dbget
'			IF rsget.EOF OR rsget.BOF THEN
'				rsget.Close
'				Alert_return("존재하지 않는 상품번호입니다. 확인 후 다시 입력해주세요")
'		       dbget.close()	:	response.End
'			End IF
'		rsget.Close
'		END IF
'
'		strSql = ""
'		strSql = " UPDATE [db_shop].[dbo].[tbl_giftkind_off] set [giftkind_name] ='"&giftkind_name&"'" +vbcrlf
'		strSql = strSql & " , [itemgubun] ='"&itemgubun&"', [itemoption] ='"&itemoption&"', [itemid] ="&itemid&"" +vbcrlf
'		strSql = strSql & " WHERE giftkind_code = "&giftkind_code
'
'		response.write strSql &"<Br>"
'		dbget.execute strSql
'
'		IF Err.Number <> 0 THEN
'			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
'	       dbget.close()	:	response.End
'		END IF
'
'		response.redirect("popgiftkindReg.asp?giftkind_name="&giftkind_name)
'		dbget.close()	:	response.End
'
'	'//신규
'	else
'
'		IF itemid = "" THEN
'		strSql = "SELECT shopitemid FROM [db_shop].dbo.tbl_shop_item where" + vbcrlf
'		strSql = strSql & " shopitemid = "&itemid&" and itemgubun = '"&itemgubun&"'" + vbcrlf
'		strSql = strSql & " and itemoption = '"&itemoption&"'" + vbcrlf
'
'		rsget.Open strSql, dbget
'			IF rsget.EOF OR rsget.BOF THEN
'				rsget.Close
'				Alert_return("존재하지 않는 상품번호입니다. 확인 후 다시 입력해주세요")
'		       dbget.close()	:	response.End
'			End IF
'		rsget.Close
'		END IF
'
'		strSql = ""
'		strSql = "INSERT INTO [db_shop].[dbo].[tbl_giftkind_off] ( [giftkind_name], [itemid] "&_
'				" ,itemgubun,itemoption) VALUES ('"&giftkind_name&"',"&itemid&",'"&itemgubun&"','"&itemoption&"') "
'		dbget.execute strSql
'
'		IF Err.Number <> 0 THEN
'			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
'	       dbget.close()	:	response.End
'		END IF
'
'		strSql = "SELECT SCOPE_IDENTITY()"
'		rsget.Open strSql, dbget
'		IF not rsget.EOF THEN
'			giftkind_code = rsget(0)
'		End IF
'		rsget.Close
'
'		response.write "<script language='javascript'>"
'		response.write "	opener.document.all.giftkind_code.value = '"&giftkind_code&"';"
'		response.write "	opener.document.all.giftkind_name.value= '"&giftkind_name&"';"
'		response.write "	window.close();"
'		response.write "</script>"
'		dbget.close()	:	response.End
'	end if

'//사은품 등록
elseif mode = "giftedit" then

	'//신규등록
	if gift_code = "" then

		'//오픈예정일 경우에는
		IF gift_status = "7" THEN
			if opendate = "" then
				 opendate = "getdate()"
			else
				opendate = " convert(nvarchar(10),'"&opendate&"',21)"&"+' "&formatdatetime(opendate,4)&"'"
			end if
		'//종료일 경우
		ELSEIF 	gift_status = "9" THEN
			if closedate = "" then
				 closedate = "getdate()"
			else
				closedate = " convert(nvarchar(10),'"&closedate&"',21)"&"+' "&formatdatetime(closedate,4)&"'"
			end if
		ELSE
			IF opendate = "" THEN
				opendate = "null"
			ELSE
				opendate = " convert(nvarchar(10),'"&opendate&"',21)"&"+' "&formatdatetime(opendate,4)&"'"
			END IF

			IF closedate = "" THEN
				closedate = "null"
			ELSE
				closedate = " convert(nvarchar(10),'"&closedate&"',21)"&"+' "&formatdatetime(closedate,4)&"'"
			END IF
		END IF

		IF giftkind_givecnt = "" THEN giftkind_givecnt = 1
		IF giftkind_limit ="" THEN giftkind_limit = 0
		IF gift_type = "" THEN gift_type =0
		IF gift_range1 = "" THEN gift_range1 = 0
		IF gift_range2 = "" THEN gift_range2 = 0

		IF giftkind_limit_sold = "" THEN giftkind_limit_sold = 0

		IF shopitemid = "" THEN shopitemid = 0
		IF itemgubun = "" THEN itemgubun = "00"
		IF itemoption = "" THEN itemoption = "0000"

		'//데이터 등록
		strSql = "INSERT INTO [db_shop].[dbo].[tbl_gift_off] " + vbcrlf
		strSql = strSql & " ([gift_name], [gift_scope], [gift_scope_add], [evt_code], [gift_type], [gift_range1], [gift_range2]" + vbcrlf
		strSql = strSql & " , makerid, [giftkind_code], gift_shopitemid, gift_itemgubun, gift_itemoption, gift_itemname, shopitemid, itemgubun, itemoption, [giftkind_type], [giftkind_cnt], [giftkind_limit], [giftkind_limit_sold], [gift_startdate]" + vbcrlf
		strSql = strSql & " , [gift_enddate],[gift_status],[adminid],opendate,lastupdate)" + vbcrlf
		strSql = strSql & " VALUES ('"&gift_name&"',"&gift_scope&", '" & gift_scope_add & "', '"&evt_code&"',"&gift_type&","&gift_range1&","&gift_range2&"" + vbcrlf
		strSql = strSql & " , '" & makerid & "',"&giftkind_code&","&gift_shopitemid&",'"&gift_itemgubun&"','"&gift_itemoption&"','" & gift_itemname & "',"&shopitemid&",'"&itemgubun&"','"&itemoption&"',"&giftkind_givecnt&","&giftkind_cnt&","&giftkind_limit&", " & giftkind_limit_sold & ",'"&gift_startdate&"','"&gift_enddate&"'" + vbcrlf
		strSql = strSql & " ,"&gift_status&",'"&session("ssBctId")&"',"&opendate&",getdate())" + vbcrlf

		'response.write strSql &"<br>"
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요")
	       dbget.close()	:	response.End
		END IF

		response.redirect("giftList.asp?menupos="&menupos&"&evt_code="&evt_code)
		dbget.close()	:	response.End

	'//수정
	else

		IF gift_status ="7" AND opendate="" THEN
			strAdd = " , [opendate] = getdate()"
		ELSEIF (gift_status = "9" and closedate ="" ) THEN
			strAdd = ", [closedate] = getdate() "	'종료처리일 설정
		END IF

		'종료일 이전에 종료시 종료일 현재 날짜로 변경
		IF gift_status = 9 and  datediff("d",gift_enddate,date()) <0 THEN
			gift_enddate = date()
		END IF

		IF giftkind_givecnt = "" THEN giftkind_givecnt = 1
		IF giftkind_limit ="" THEN giftkind_limit = 0
		IF gift_type = "" THEN gift_type =0
		IF gift_range1 = "" THEN gift_range1 = 0
		IF gift_range2 = "" THEN gift_range2 = 0

		IF giftkind_limit_sold = "" THEN giftkind_limit_sold = 0

		IF shopitemid = "" THEN shopitemid = 0
		IF itemgubun = "" THEN itemgubun = "00"
		IF itemoption = "" THEN itemoption = "0000"

	 	'//데이터 수정
		strSql = " UPDATE [db_shop].[dbo].[tbl_gift_off] SET  [gift_name] = '"&gift_name&"', [gift_scope]="&gift_scope&_
				" , [gift_type]="&gift_type&", [gift_range1]="&gift_range1&", [gift_range2]= "&gift_range2&_
				" , makerid = '" & makerid & "', [giftkind_code]= "&giftkind_code&", [gift_shopitemid]= "&gift_shopitemid&", [gift_itemgubun]= '"&gift_itemgubun&"', [gift_itemoption]= '"&gift_itemoption&"', gift_itemname = '" & gift_itemname & "', [shopitemid]= "&shopitemid&", [itemgubun]= '"&itemgubun&"', [itemoption]= '"&itemoption&"', [giftkind_type] ="&giftkind_givecnt&" , [giftkind_cnt]= "&giftkind_cnt&_
				" , [giftkind_limit]="&giftkind_limit&", [giftkind_limit_sold]="&giftkind_limit_sold&" , [gift_startdate]= '"&gift_startdate&"', [gift_enddate]='"&gift_enddate&"'"&_
				" , [gift_status] = "&gift_status&", [gift_using] = '"&gift_using&"'"&_
				" , [adminid]= '"&session("ssBctId")&"', [lastupdate] = getdate() "&strAdd&_
				" WHERE gift_code = "&gift_code

		'response.write strSql
		dbget.execute strSql

		IF Err.Number <> 0 THEN
			Alert_return("데이터 처리에 문제가 발생하였습니다.관리자에게 문의해 주세요1")
	       dbget.close()	:	response.End
		END IF

		response.redirect("giftList.asp?evt_code="&evt_code&"&menupos="&menupos)
		dbget.close()	:	response.End

	end if

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->