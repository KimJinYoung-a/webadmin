<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' PageName : index.asp
' Discription : 굿노트 다이어리 스티커 컨텐츠 등록 창
' History : 2023.03.31 정태훈
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/eventmanage/event/v5/lib/admineventhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function_v5.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/event/goodNoteDiaryCls.asp"-->
<%
dim menupos, iCurrpage, iPageSize, iPerCnt, iTotCnt
dim arrList, cEvtList, ix
	menupos = request("menupos")
	iCurrpage = Request("iC")	'현재 페이지 번호
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 20		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격

	'데이터 가져오기
	set cEvtList = new GoodNoteDiaryCls
		cEvtList.FCurrPage = iCurrpage	'현재페이지
		cEvtList.FPageSize = iPageSize '한페이지에 보이는 레코드갯수
 		arrList = cEvtList.getStickerList	'데이터목록 가져오기
 		iTotCnt = cEvtList.FTotalCount	'전체 데이터  수
 	set cEvtList = nothing
%>
<script>
function TnTrainThemeItemBannerReg(){
    var winpop = window.open("/admin/eventmanage/goodnote/pop_sticker_register.asp","winpop","width=1200,height=800,scrollbars=yes,resizable=yes");
    winpop.focus();
}
function fnStickerEdit(idx){
    var winEditpop = window.open("/admin/eventmanage/goodnote/pop_sticker_register.asp?idx="+idx,"winpop","width=1200,height=800,scrollbars=yes,resizable=yes");
    winEditpop.focus();
}
</script>
<div class="popV19">
	<div class="popHeadV19">
		<h1>굿노트 다이어리 스티커 관리</h1>
	</div>
    <button class="btn4 btnBlock btnWhite2 tMar10 tPad20 bPad20 lt" onClick="TnTrainThemeItemBannerReg();return false;"><span class="mdi mdi-plus cBl4 fs15"></span> 스티커 추가</button>
    <% If isArray(arrList) Then %>
    <div class="tableV19BWrap tMar15 tPad25 topLineGrey2">
        <table class="tableV19A tableV19B tMar10">
            <thead>
                <tr>
                    <th>No</th>
                    <th>제목</th>
                    <th>오픈일</th>
                    <th>사용여부</th>
                    <th>등록일</th>
                </tr>
            <thead>
            <tbody>
                <% For ix = 0 To UBound(arrList,2) %>
                <tr onclick="fnStickerEdit(<%=arrList(0,ix)%>);">
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="mdi fs20"><%=arrList(0,ix)%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="mdi fs20"><%=arrList(1,ix)%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><span class="previewThumb50W"><%=FormatDate(arrList(2,ix),"0000.00.00")%>~<%=FormatDate(arrList(3,ix),"0000.00.00 00:00:00")%></span></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><%=arrList(4,ix)%></td>
                    <td<% if arrList(4,ix)="N" then response.write " style='background-color:#ebebeb;'" %>><%=FormatDate(arrList(5,ix),"0000.00.00 00:00:00")%></td>
                </tr>
                <% Next %>
            </tbody>
        </table>
    </div>
    <% End If %>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->