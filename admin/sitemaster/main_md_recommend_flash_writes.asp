<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
'###############################################
' PageName : main_md_recommand_flash_writes.asp
' Discription : 상품코드 일괄 등록
'###############################################
'// 변수 선언
Dim realdate : realdate = request("realdate")
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
// 폼검사
function SaveForm(frm) {
	var selChk=true;
	if(frm.linkitemid.value=="") {
		alert("일괄 등록하실 상품코드를 입력해주세요");
		frm.linkitemid.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<form name="frmSub" method="post" action="main_md_recommend_flash_proc.asp" style="margin:0px;">
    <table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
        <tr bgcolor="#FFFFFF">
            <td height="25" colspan="4" bgcolor="#F8F8F8"><b>소재 정보 - 상품코드 일괄 등록</b></td>
        </tr>
        <colgroup>
            <col width="100" />
            <col width="*" />
            <col width="100" />
            <col width="*" />
        </colgroup>
        <tr bgcolor="#FFFFFF">
            <td bgcolor="#DDDDFF">상품코드</td>
            <td colspan="3">
                <textarea name="linkitemid" class="textarea" title="상품코드" style="width:95%; height:80px;"></textarea>
                <p>※ 상품코드를 쉼표(,) 또는 엔터로 구분하여 입력</p>
                <p>※ 상품명은 기본 상품명으로 입력 됩니다. (수정 필요)</p>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td width="150" bgcolor="#DDDDFF">반영시작일</td>
            <td colspan="3">
                <input id="startdate" name="startdate" value="<%=Left(realdate,10)%>" class="text" size="10" maxlength="10" />
                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer;" align="absbottom" />
                <input type="text" name="startdatetime" size="2" maxlength="2" value="00" />(시 00~23)
                <input type="text" name="dummy0" value="00:00" size="6" readonly class="text_ro" />
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
        <tr bgcolor="#FFFFFF">
            <td width="150" bgcolor="#DDDDFF">반영종료일</td>
            <td colspan="3">
                <input id="enddate" name="enddate" value="<%=Left(realdate,10)%>" class="text" size="10" maxlength="10" />
                <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absbottom" />
                <input type="text" name="enddatetime" size="2" maxlength="2" value="23">(시 00~23)
                <input type="text" name="dummy1" value="59:59" size="6" readonly class="text_ro" />
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
        <tr bgcolor="#FFFFFF">
            <td bgcolor="#DDDDFF">전시순서</td>
            <td>
                <input type="text" name="disporder" class="text" size="4" value="99" />
            </td>
            <td bgcolor="#DDDDFF">사용여부</td>
            <td>
                <span id="rdoUsing">
                <input type="radio" name="isusing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">사용</label>
                <input type="radio" name="isusing" id="rdoUsing2" value="N" /><label for="rdoUsing2">삭제</label>
                </span>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveForm(this.form);"></td>
        </tr>
    </table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->