<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Event_cls.asp"-->
<%
	'/// 문서형태를 Ms-Excel로 지정 ///
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition","Attachment;"

	'// 변수 선언 //
	dim evtId

	dim oPart, lp

	'// 파라메터 접수 //
	evtId = RequestCheckvar(request("evtId"),10)

	'// 클래스 선언
	set oPart = new CPart
	oPart.FRectevtId = evtId

	oPart.GetPartAllList


	'// 필드값 출력 //
	Response.write ("이벤트 번호 : " & chr(9))
	Response.write (evtId & chr(9))
	Response.write (chr(9))
	Response.write ("참여자 수 :" & chr(9))
	Response.write (oPart.FTotalCount & chr(9))
	Response.Write (chr(13) & chr(10))
	Response.Write (chr(13) & chr(10))

	Response.write ("번호" & chr(9))
	Response.write ("아이디" & chr(9))
	Response.write ("회원등급" & chr(9))
	Response.write ("이름" & chr(9))
	Response.write ("내용1" & chr(9))
	Response.write ("내용2" & chr(9))
	Response.write ("참여수" & chr(9))
	Response.write ("참여일시" & chr(9))
	Response.write ("구매액(6M)" & chr(9))
	Response.write ("회원가입일" & chr(9))
	Response.write ("당첨횟수" & chr(9))
	Response.Write (chr(13) & chr(10))

	if oPart.FTotalCount>0 then

	'@@ 도돌이 시작
	for lp=0 to oPart.FTotalCount - 1

		Response.write (oPart.FPartList(lp).FprtId & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserId & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserLevel & chr(9))
		Response.write (oPart.FPartList(lp).FprtUserNm & chr(9))
		Response.write (Replace(db2html(oPart.FPartList(lp).FprtCont1), chr(13)&chr(10), " ") & chr(9))
		Response.write (Replace(db2html(oPart.FPartList(lp).FprtCont2), chr(13)&chr(10), " ") & chr(9))
		Response.write (oPart.FPartList(lp).FprtCnt & chr(9))
		Response.write (oPart.FPartList(lp).FprtDate & chr(9))
		Response.write (FormatNumber(oPart.FPartList(lp).FsixMonthOrder,0) & chr(9))
		Response.write (oPart.FPartList(lp).FregDate & chr(9))
		Response.write (oPart.FPartList(lp).FprizeCnt & chr(9))
		Response.Write (chr(13) & chr(10))

	next

	end if

set oPart = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->