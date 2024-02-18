<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 장비자산관리 공통코드
' History : 2008년 06월 27일 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim gubuntype, gubuncd, typename, gubunname, isusing, orderno ,mode , strSql , idx
	gubuntype		= requestCheckVar(Request("gubuntype"),10)
	gubuncd		= requestCheckVar(Request("gubuncd"),2)
	typename		= requestCheckVar(Request("typename"),64)
	gubunname		= requestCheckVar(Request("gubunname"),64)
	isusing		= requestCheckVar(Request("isusing"),1)
	orderno		= requestCheckVar(Request("orderno"),10)
	mode		= requestCheckVar(Request("mode"),64)
	idx		= requestCheckVar(Request("idx"),64)		

'/신규등록
IF mode = "I" THEN
	strSql = "SELECT *"
	strSql = strSql & " FROM db_partner.dbo.tbl_equipment_gubun"
	strSql = strSql & " Where gubuntype="&gubuntype&" and gubuncd='"&gubuncd&"'"
	
	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>"
		response.write "	alert('이미존재하는 코드값 입니다.'); history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO db_partner.dbo.tbl_equipment_gubun (gubuntype, gubuncd, typename, gubunname, isusing, orderno)"&_
			" Values("&gubuntype&",'"&gubuncd&"','"&html2db(getequipmentCodeType(trim(gubuntype)))&"','"&gubunname&"','"&isusing&"',"&orderno&") "
	
	'response.write strSql &"<br>"
	dbget.execute strSql

	response.write "<script>alert('ok'); location.href='/common/equipment/popmanagecode.asp?gubuntype="&gubuntype&"';</script>"
	
'/수정
ELSEIF mode="U" THEN

	strSql = "SELECT *"
	strSql = strSql & " FROM db_partner.dbo.tbl_equipment_gubun"
	strSql = strSql & " Where gubuntype="&gubuntype&" and gubuncd='"&gubuncd&"'"
	strSql = strSql & " and idx <> "&idx&""
	
	'response.write strSql &"<Br>"
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		response.write "<script>"
		response.write "	alert('이미존재하는 코드값 입니다.'); history.go(-1);"
		response.write "</script>"
		dbget.close()	:	response.End
	end if
	rsget.close
	
	strSql =" UPDATE db_partner.dbo.tbl_equipment_gubun"&_
			" Set gubuntype = '"&gubuntype&"'"&_
			" ,gubuncd = '"&gubuncd&"'"&_
			" ,typename = '"&html2db(getequipmentCodeType(trim(gubuntype)))&"'"&_
			" ,gubunname = '"&gubunname&"'"&_
			" ,isusing = '"&isusing&"'"&_
			" ,orderno = '"&orderno&"'"&_
			" WHERE idx ='"&idx&"'"
	
	'response.write strSql &"<br>"
	dbget.execute strSql

	response.write "<script>alert('ok'); location.href='/common/equipment/popmanagecode.asp?gubuntype="&gubuntype&"';</script>"
END IF	
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->