<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
Dim eCode, ename, enameEng, cEvtCont
	eCode		= requestCheckVar(Request("eC"),10)	'이벤트코드

	set cEvtCont = new ClsEvent
	cEvtCont.FECode = eCode	'이벤트 코드
	'이벤트 내용 가져오기
	cEvtCont.fnGetEventCont

	ename 		=	db2html(cEvtCont.FEName)
	enameEng =	db2html(cEvtCont.FENameEng) '이벤트 영문 추가

%>

<script type="text/javascript">
<!--
//-- jsEvtSubmit(form 명) : 이벤트 수정처리 --//
	function jsEvtSubmit(frm){
alert("a");
return false;
	}
//-->
</script>
<form name="frmEvt" method="post" action="event_process.asp" onSubmit="return jsEvtSubmit(this);" style="margin:0px;">
		   	<%
		   		'// 브랜드할인이면 제목을 할인율로 표시
		   		dim arrEname
				arrEname = Split(ename,"|")
				if Ubound(arrEname)<2 then
					arrEname = ename & "|0|0"
					arrEname = Split(arrEname,"|")
				end if

				If enameEng <> "" then
					Dim arrEnameEng
					arrEnameEng = Split(enameEng,"|")
					if Ubound(arrEnameEng)<2 then
						arrEnameEng = enameEng & "|0|0"
						arrEnameEng = Split(arrEnameEng,"|")
					end If
				End If
		   	%>
					이벤트명: <input type="text" name="sEDN" size="50" maxlength="50" value="<%=arrEname(0)%>"><br>
		<input type="image" src="/images/icon_save.gif">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->