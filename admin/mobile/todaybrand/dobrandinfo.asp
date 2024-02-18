<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'// 변수 선언 및 전달값 저장
Dim mode, mainimg, moreimg
Dim isusing, idx
Dim linkurl, sqlStr
Dim adminid, ordertext
Dim startdate, enddate
Dim itemid1, itemid2, evt_code
Dim makerid, maincopy, subcopy, menupos

	mode = request.form("mode")
	idx	= request.form("idx")
	startdate = request.form("StartDate")& " " &request.form("sTm")
	enddate	= request.form("EndDate")& " " &request.form("eTm")
	isusing	= request.form("isusing")
	ordertext = request.form("ordertext")
	adminid	= request.form("adminid") '로그인 아이디
	makerid	= Trim(request.form("makerid")) '브랜드 ID
	maincopy = Trim(request.form("maincopy")) '메인카피
	subcopy	= request.form("subcopy") '서브카피
	linkurl	= Trim(request.form("linkurl")) '링크 주소 default 브랜드 + 기타 4가지 사용 가능
	itemid1	= Trim(request.form("itemid1")) '2017-08-03 itemid1
	itemid2	= Trim(request.form("itemid2")) '2017-08-03 itemid2
    mainimg	= request.form("mainimg")
    moreimg	= request.form("moreimg")
    menupos	= request.form("menupos")
if mode="" then
    Call Alert_return("not valid code.")
    dbget.Close: Response.End   
end if

If mode = "add" Then '//날짜 체크
	Dim itemcount
	sqlStr = "select count(*) from db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] where startdate = '"& startdate &"' and isusing = 'Y'" 
	rsget.Open SqlStr, dbget, 1
	if Not rsget.Eof Then
		itemcount = rsget(0)
	end If
	rsget.Close

	If itemcount > 0 Then 
		'response.write "<script>alert('동일날짜 혹은 동일시간 대에 시작 하는 컨텐츠가 있습니다.');self.location.href='/admin/mobile/todaybrand/';</script>"
        response.write "<script>alert('동일날짜 혹은 동일시간 대에 시작 하는 컨텐츠가 있습니다.');history.go(-1);</script>"
		Response.end
	End If 
End If

if instr(linkurl,"/event/eventmain.asp?eventid=")>0 then
    evt_code=replace(linkurl,"/event/eventmain.asp?eventid=","")
else
    evt_code=0
end if

'/신규등록
if mode="add" then
	'// 신규 저장
    sqlStr = " insert into db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] " & VbCrlf
    sqlStr = sqlStr & " (startdate,enddate,adminid,isusing,ordertext,makerid,maincopy,subcopy,linkurl,itemid1,itemid2,mainimg,moreimg,eventcode) " & VbCrlf
    sqlStr = sqlStr & " values(" & VbCrlf
    sqlStr = sqlStr & " '" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & enddate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & adminid & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & isusing & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & ordertext & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & makerid & "'" & VbCrlf
    sqlStr = sqlStr & " ,N'" & html2db(maincopy) & "'" & VbCrlf
    sqlStr = sqlStr & " ,N'" & html2db(subcopy) & "'" & VbCrlf
    sqlStr = sqlStr & " ,N'" & linkurl & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & itemid1 & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & itemid2 & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & mainimg & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & moreimg & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & evt_code & "'" & VbCrlf
    sqlStr = sqlStr & " )"
'	response.write sqlStr &"<Br>"
'	Response.end
    dbget.Execute sqlStr

'// 수정

'makerid,maincopy,subcopy,linkurl,itemid1,itemid2

else
   sqlStr = " update db_sitemaster.[dbo].[tbl_mobile_main_brandinfo] " & VbCrlf
   sqlStr = sqlStr & " set startdate='" & startdate & "'" & VbCrlf
   sqlStr = sqlStr & " , enddate ='" & enddate & "'" & VbCrlf
   sqlStr = sqlStr & " , lastadminid='" & adminid & "'" & VbCrlf
   sqlStr = sqlStr & " , isusing='" & isusing & "'" & VbCrlf
   sqlStr = sqlStr & " , ordertext='" & ordertext & "'" & VbCrlf
   sqlStr = sqlStr & " , lastupdate=getdate()" & VbCrlf
   sqlStr = sqlStr & " , makerid='" & makerid & "'" & VbCrlf
   sqlStr = sqlStr & " , maincopy='" & html2db(maincopy) & "'" & VbCrlf
   sqlStr = sqlStr & " , subcopy='" & html2db(subcopy) & "'" & VbCrlf
   sqlStr = sqlStr & " , linkurl='" & linkurl & "'" & VbCrlf
   sqlStr = sqlStr & " , itemid1='" & itemid1 & "'" & VbCrlf
   sqlStr = sqlStr & " , itemid2='" & itemid2 & "'" & VbCrlf
   sqlStr = sqlStr & " , mainimg='" & mainimg & "'" & VbCrlf
   sqlStr = sqlStr & " , moreimg='" & moreimg & "'" & VbCrlf
   sqlStr = sqlStr & " , eventcode='" & evt_code & "'" & VbCrlf
   sqlStr = sqlStr & " where idx='" & Cstr(idx) & "'"
   
   'response.write sqlStr &"<Br>"
   'response.end
   dbget.Execute sqlStr
end if

dim referer
referer = request.ServerVariables("HTTP_REFERER")
response.write "<script>alert('저장되었습니다.');</script>"
response.write "<script>location.replace('/admin/mobile/todaybrand/?menupos=" + Cstr(menupos) + "');</script>"
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->