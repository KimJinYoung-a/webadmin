<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기획전/이벤트 링크 관리 등록페이지
' History : 2022-08-08 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/exhibitionitems/lib/classes/exhibitionCls.asp"-->
<%
    dim mastercode : mastercode = request("mastercode")
    dim idx : idx = request("idx")
    dim oExhibition
    dim mode

    if idx = 0 then 
        mode = "evtlinkadd"
    else
        mode = "evtlinkmodify"
    end if 

    set oExhibition = new ExhibitionCls
        oExhibition.Frectidx = idx
        oExhibition.getOneEventLinkContents()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function fnEventSave(frm){
    if (!frm.evt_code.value) {
        alert("이벤트 코드를 입력 해주세요.");
        frm.evt_code.focus;
    }

    if (!frm.StartDate.value) {
        alert("배너 노출 시작일을 입력 해주세요.");
        frm.StartDate.focus;
    }

    if (!frm.EndDate.value) {
        alert("배너 노출 종료일을 입력 해주세요.");
        frm.EndDate.focus;
    }
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}

function jsLastEvent(num){
    winLast = window.open('pop_event_lastlist.asp?num='+num,'pLast','width=800,height=600, scrollbars=yes')
    winLast.focus();
}
</script>
</head>
<body>
<div class="contSectFix scrl">
	<div class="pad20">
		<form name="frmreg" method="post" action="/admin/exhibitionitems/lib/exhibition_proc.asp">
        <input type="hidden" name="mode" value="<%=mode%>"/>
        <input type="hidden" name="eidx" value="<%=idx%>"/>
        <input type="hidden" name="mastercode" value="<%=mastercode%>">
		<table class="tbType1 listTb">
			<tr>
				<td>
					<table class="tbType1 listTb">
						<tr bgcolor="#FFFFFF" height="25">
							<td colspan="2" ><b>이벤트 등록</b></td>
						</tr>
						<tr  bgcolor="<%= adminColor("tabletop") %>">
							<th> 이벤트 코드</th>
							<td bgcolor="#FFFFFF" style="text-align:left;">
								<input type="text" size="10" name="evt_code" value="<%=oExhibition.FOneItem.Fevt_code%>"/> <input type="button" value="이벤트 불러오기" onclick="jsLastEvent(1);"/>
                                <div class="tPad15" id="infomenu" style="display:<%=chkiif(idx > 0 , "", "none")%>;">
                                    <div>이벤트명 : <span id="evt_name"><%=oExhibition.FOneItem.Fevt_name%></span></div>
                                    <div>시작일 : <span id="evt_startdate"><%=oExhibition.FOneItem.Fevt_startdate%></span></div>
                                    <div>종료일 : <span id="evt_enddate"><%=oExhibition.FOneItem.Fevt_enddate%></span></div>
                                    <div>할인율 : <span id="evt_saleper" style='color:red'><%=chkiif(oExhibition.FOneItem.Fsaleper <> "",oExhibition.FOneItem.Fsaleper,"할인 정보가 없습니다.")%></span></div>
                                    <div>쿠폰할인 : <span id="evt_salecoupon" style='color:green'><%=chkiif(oExhibition.FOneItem.Fsalecper <> "",oExhibition.FOneItem.Fsalecper,"쿠폰할인 정보가 없습니다.")%></span></div>
                                </div>
							</td>
						</tr>
                        <tr>
                            <th>링크 제목</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="title" size="70" value="<%=oExhibition.FOneItem.Ftitle%>">
                            </td>
                        </tr>
                        <tr>
                            <th>시작일</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="StartDate" id="startdate" value="<%=oExhibition.FOneItem.Fstartdate%>">
                                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" style="vertical-align:middle;"/>
                                <script type="text/javascript">
                                var CAL_Start = new Calendar({
                                    inputField : "startdate",
                                    trigger    : "startdate_trigger",
                                    onSelect: function() {
                                        var date = Calendar.intToDate(this.selection.get());
                                        CAL_End.args.min = date;
                                        CAL_End.redraw();
                                        this.hide();
                                    },
                                    bottomBar: true,
                                    dateFormat: "%Y-%m-%d"
                                });
                                </script>
                            </td>
                        </tr>
                        <tr>
                            <th>종료일</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="EndDate" id="enddate" value="<%=oExhibition.FOneItem.Fenddate%>">
                                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" style="vertical-align:middle;"/>
                                <script type="text/javascript">
                                var CAL_End = new Calendar({
                                    inputField : "enddate",
                                    trigger    : "enddate_trigger",
                                    onSelect: function() {
                                        var date = Calendar.intToDate(this.selection.get());
                                        CAL_Start.args.max = date;
                                        CAL_Start.redraw();
                                        this.hide();
                                    },
                                    bottomBar: true,
                                    dateFormat: "%Y-%m-%d"
                                });
                                </script>
                            </td>
                        </tr>
                        <tr>
                            <th>우선순위</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="text" name="evtsorting" value="<%=chkiif(oExhibition.FOneItem.Fevtsorting = "","99",oExhibition.FOneItem.Fevtsorting)%>">
                            </td>
                        </tr>
                        <tr>
                            <th>사용여부</th>
                            <td bgcolor="#FFFFFF" style="text-align:left;">
                                <input type="radio" name="evtisusing" value="1" id="usey" <%=chkiif(oExhibition.FOneItem.Fisusing = ""  or oExhibition.FOneItem.Fisusing = "1" , "checked" , "")%>> <label for="usey">사용함</label>
                                <input type="radio" name="evtisusing" value="0" id="usen" <%=chkiif(oExhibition.FOneItem.Fisusing = "0" , "checked" , "")%>> <label for="usen">사용안함</label>
                            </td>
                        </tr>
					</table>
				</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td colspan="2">
					<img src="http://webadmin.10x10.co.kr/images/icon_save.gif" border="0" onClick="fnEventSave(frmreg);" style="cursor:pointer">
					<img src="http://webadmin.10x10.co.kr/images/icon_cancel.gif" border="0" onClick="window.close();" style="cursor:pointer">
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
<%
    set oExhibition = nothing
%>
<!-- 리스트 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->