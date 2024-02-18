<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : domdpick.asp
' Discription : mdpick 처리 페이지
' History : 2013.12.16 이종화 생성
'###############################################

'// 변수 선언 및 파라메터 접수
dim menupos, mode
Dim idx , eventid , linkurl , maincopy , subcopy , startdate , enddate , prevDate
Dim evtstdate , evteddate , isusing , ordertext , sortnum , preDate , adminid , sale_per , coupon_per , tag_only, dispOption, contentType, contentImg, itemId, pcwebIsUsing, mobileIsUsing, event_info_option, event_info
Dim sqlStr

	menupos				= Request("menupos")
	mode				= request("mode")
	idx					= request("idx")
	eventid				= request("eventid")
	linkurl				= request("linkurl")
	maincopy			= request("maincopy")
	subcopy				= request("subcopy")
	startdate			= request("StartDate")& " " &request("sTm")
	enddate				= request("EndDate")& " " &request("eTm")
	evtstdate			= Left(request("evtstdate"),10)
	evteddate			= Left(request("evteddate"),10)
	isusing				= request("isusing")
	ordertext			= request("ordertext")
	sortnum				= request("sortnum")
	prevDate			= request("prevDate")
	adminid				= request("adminid") '로그인 아이디
	sale_per			= request("sale_per") '2017-07-27 세일 text
	coupon_per			= request("coupon_per") '2017-07-27 쿠폰 text
    tag_only			= request("tag_only") '2018-08-08 태그
    dispOption          = request("dispOption")
    contentType         = request("contentType")
    contentImg          = request("contentImg")
    itemId              = request("itemId")
    pcwebIsUsing        = request("PCisusing")
    mobileIsUsing       = request("mobileisusing")
    event_info_option   = request("event_info_option")
    event_info          = request("event_info")

'// 모드에 따른 분기
'/신규등록
if mode="add" then
	'// 신규 저장
    sqlStr = " INSERT INTO db_sitemaster.dbo.tbl_pcmain_enjoyevent " & VbCrlf
    sqlStr = sqlStr & " ( linkurl , maincopy, evtstdate, evteddate , startdate , enddate , adminid , isusing , ordertext , sortnum , eventid , subcopy , sale_per , coupon_per , tag_only, dispOption, contentType, contentImg, itemId, event_info, event_info_option) " & VbCrlf
    sqlStr = sqlStr & " VALUES (" & VbCrlf
    sqlStr = sqlStr & " '" & linkurl & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & maincopy & "'" & VbCrlf
    if contentType = "2" and dispOption = "2" then
    sqlStr = sqlStr & " ,'" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & enddate & "'" & VbCrlf    
    else 
    sqlStr = sqlStr & " ,'" & evtstdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & evteddate & "'" & VbCrlf    
    end if
    sqlStr = sqlStr & " ,'" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & enddate & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & adminid & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & isusing & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & ordertext & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & sortnum & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & eventid & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & subcopy & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & sale_per & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & coupon_per & "'" & VbCrlf
    sqlStr = sqlStr & " ,'" & tag_only & "'" & VbCrlf    
    sqlStr = sqlStr & " ,'" & dispOption & "'" & VbCrlf     
    sqlStr = sqlStr & " ,'" & contentType & "'" & VbCrlf            
    sqlStr = sqlStr & " ,'" & contentImg & "'" & VbCrlf            
    sqlStr = sqlStr & " ,'" & itemId & "'" & VbCrlf            
    sqlStr = sqlStr & " ,'" & event_info & "'" & VbCrlf            
    sqlStr = sqlStr & " ,'" & event_info_option & "'" & VbCrlf            
    sqlStr = sqlStr & " )"

'	response.write sqlStr &"<Br>"
'	Response.end
    dbget.Execute sqlStr

'// 수정
elseif mode="exhibitionOpenCtrl" then
    sqlStr = " UPDATE db_sitemaster.dbo.tbl_pcmain_top_exhibition_ctrl " & VbCrlf
    sqlStr = sqlStr & " SET isUsing='" & pcwebIsUsing & "'" & VbCrlf    
    sqlStr = sqlStr & " WHERE flatform ='PCWEB' "
    sqlStr = sqlStr & " UPDATE db_sitemaster.dbo.tbl_pcmain_top_exhibition_ctrl " & VbCrlf
    sqlStr = sqlStr & " SET isUsing='" & mobileIsUsing & "'" & VbCrlf    
    sqlStr = sqlStr & " WHERE flatform ='MOBILE' "
   'response.write sqlStr &"<Br>"
   'response.end
   dbget.Execute sqlStr
else
    sqlStr = " UPDATE db_sitemaster.dbo.tbl_pcmain_enjoyevent " & VbCrlf
    sqlStr = sqlStr & " SET linkurl='" & linkurl & "'" & VbCrlf
    sqlStr = sqlStr & " , maincopy='" & maincopy & "'" & VbCrlf
    if contentType = "2" and dispOption = "2" then
    sqlStr = sqlStr & " , evtstdate='" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " , evteddate='" & enddate & "'" & VbCrlf
    else 
    sqlStr = sqlStr & " , evtstdate='" & evtstdate & "'" & VbCrlf
    sqlStr = sqlStr & " , evteddate='" & evteddate & "'" & VbCrlf
    end if    
    sqlStr = sqlStr & " , startdate='" & startdate & "'" & VbCrlf
    sqlStr = sqlStr & " , enddate ='" & enddate & "'" & VbCrlf
    sqlStr = sqlStr & " , lastadminid='" & adminid & "'" & VbCrlf
    sqlStr = sqlStr & " , isusing='" & isusing & "'" & VbCrlf
    sqlStr = sqlStr & " , ordertext='" & ordertext & "'" & VbCrlf
    sqlStr = sqlStr & " , sortnum='" & sortnum & "'" & VbCrlf
    sqlStr = sqlStr & " , eventid='" & eventid & "'" & VbCrlf
    sqlStr = sqlStr & " , lastupdate=getdate()" & VbCrlf
    sqlStr = sqlStr & " , subcopy='" & subcopy & "'" & VbCrlf
    sqlStr = sqlStr & " , sale_per='" & sale_per & "'" & VbCrlf
    sqlStr = sqlStr & " , coupon_per='" & coupon_per & "'" & VbCrlf
    sqlStr = sqlStr & " , tag_only='" & tag_only & "'" & VbCrlf   
    sqlStr = sqlStr & " , dispOption='" & dispOption & "'" & VbCrlf   
    sqlStr = sqlStr & " , contentType='" & contentType & "'" & VbCrlf   
    sqlStr = sqlStr & " , contentImg='" & contentImg & "'" & VbCrlf   
    sqlStr = sqlStr & " , itemId='" & itemId & "'" & VbCrlf   
    sqlStr = sqlStr & " , event_info='" & event_info & "'" & VbCrlf   
    sqlStr = sqlStr & " , event_info_option='" & event_info_option & "'" & VbCrlf   
    sqlStr = sqlStr & " WHERE idx='" & Cstr(idx) & "'"

'   response.write sqlStr &"<Br>"
'   response.end
   dbget.Execute sqlStr
end if

%>
<% if mode = "exhibitionOpenCtrl" then %>
<script type="text/javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	window.close();
//-->
</script>
<% else %>
<script type="text/javascript">
<!--
	// 목록으로 복귀
	alert("저장했습니다.");
	self.location = "index.asp?menupos=<%=menupos%>&prevDate=<%=prevDate%>&dispOption=<%=dispOption%>";
//-->
</script>
<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
