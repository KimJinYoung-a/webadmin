<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->

<%
Dim arrList,intLoop, clsCode, sMode, searchsendmethod, menupos
dim tidx, sendmethod, template_code, template_name, contents, button_name, button_url_mobile, button_name2, button_url_mobile2
dim failed_type, failed_subject, failed_msg, isusing, regadminid, lastadminid, regdate, lastupdate, sortno
    tidx = requestcheckvar(getNumeric(request("tidx")),10)
    menupos = requestcheckvar(getNumeric(request("menupos")),10)
    searchsendmethod = requestcheckvar(request("searchsendmethod"),16)
	sMode ="I"

if searchsendmethod="" or isnull(searchsendmethod) then searchsendmethod = "KAKAOALRIM"

Set clsCode = new ClmstargetCommonCode  	
	IF tidx <> "" THEN
		sMode ="U"
		clsCode.frecttidx  = tidx 
		clsCode.GetlmstemplateCont

        if clsCode.FTotalCount>0 THEN
            tidx = clsCode.ftidx
            sendmethod = clsCode.fsendmethod
            template_code = clsCode.ftemplate_code
            template_name = clsCode.ftemplate_name
            contents = clsCode.fcontents
            button_name = clsCode.fbutton_name
            button_url_mobile = clsCode.fbutton_url_mobile
            button_name2 = clsCode.fbutton_name2
            button_url_mobile2 = clsCode.fbutton_url_mobile2
            failed_type = clsCode.ffailed_type
            failed_subject = clsCode.ffailed_subject
            failed_msg = clsCode.ffailed_msg
            isusing = clsCode.fisusing
            regadminid = clsCode.fregadminid
            lastadminid = clsCode.flastadminid
            regdate = clsCode.fregdate
            lastupdate = clsCode.flastupdate
            sortno = clsCode.fsortno
        end if
    END IF
 		 
	clsCode.frectsendmethod = searchsendmethod
	arrList = clsCode.GetlmstemplateList
Set clsCode = nothing 

if sendmethod="" or isnull(sendmethod) then sendmethod = "KAKAOALRIM"
if isusing="" or isnull(isusing) then isusing = "Y"
if sortno="" or isnull(sortno) then sortno = 0
%>
<script type='text/javascript'>

	// 코드타입 변경이동
	function jsSetCode(tidx){	
		self.location.href = "/admin/appmanage/lms/lmstemplatereg.asp?tidx="+tidx+"&menupos=<%= menupos %>";
	}
	
	//코드 검색
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;

        if (frm.sendmethod.value.length<1){
            alert('발송방법을 선택해주세요');
			frm.sendmethod.focus();
			return;
        }
        if (frm.template_code.value.length<1){
            alert('템플릿코드를 입력해 주세요');
			frm.template_code.focus();
			return;
        }
        if (frm.template_name.value.length<1){
            alert('템플릿명을 입력해 주세요');
			frm.template_name.focus();
			return;
        }
		if (frm.contents.value==''){ 
			alert('내용을 등록해 주세요.');
			frm.contents.focus();
			return;
		}
        //if (frm.button_name.value.length<1){
        //    alert('카카오톡 버튼 이름을 입력해 주세요.');
        //    frm.button_name.focus();
        //    return;
        //}
        //if (frm.button_url_mobile.value.length<1){
        //    alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');
        //    frm.button_url_mobile.focus();
        //    return;
        //}
        if (frm.failed_subject.value!=''){
            if (GetByteLength(frm.failed_subject.value) > 50){
                alert("카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.");
                frm.failed_subject.focus();
                return;
            }
        }
		if(!frm.isusing.value) {
			alert("사용여부를 입력해 주세요");
			frm.isusing.focus();
			return false;
		}
        if (frm.sortno.value.length<1){
            alert('정렬순서를 입력해 주세요.');
            frm.sortno.focus();
            return;
        }
		frm.submit();
	}

</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td>
		<form name="frmReg" method="post" action="/admin/appmanage/lms/lmstemplate_process.asp" style="margin:0px;">	
        <input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="sM" value="<%=sMode%>">
        <input type="hidden" name="tidx" value="<%= tidx %>">
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">		
        <% if tidx<>"" then %>								
            <tr>
                <td bgcolor="#EFEFEF" width="100" align="center">템플릿번호</td>
                <td bgcolor="#FFFFFF">
                    <%= tidx %>
                </td>
            </tr>
        <% end if %>
        <tr>
            <td bgcolor="#EFEFEF" align="center">발송방법</td>
            <td bgcolor="#FFFFFF">
                <% Drawsendmethod "sendmethod", sendmethod, "", "Y" %>
            </td>
        </tr>			
        <tr>
            <td bgcolor="#EFEFEF" align="center">템플릿코드</td>
            <td bgcolor="#FFFFFF">
                <input type="text" name="template_code" value="<%= template_code %>" maxlength=32 size=26>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">템플릿명</td>
            <td bgcolor="#FFFFFF">
                <input type="text" name="template_name" value="<%= template_name %>" maxlength=64 size=52>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">내용</td>
            <td bgcolor="#FFFFFF">
                <textarea name="contents" cols=100 rows=8><%= contents %></textarea>
				<br>※ 입력시 실제 고객 데이터로 치환됨.
				<br>고객ID : <font color="red">${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}</font>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">버튼이름1</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="button_name" value="<%= button_name %>" size="64" maxlength=64 />
                예) 확인하러 가기
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">버튼모바일주소1</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="button_url_mobile" value="<%= button_url_mobile %>" size="120" maxlength=256 />
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">버튼이름2</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="button_name2" value="<%= button_name2 %>" size="64" maxlength=64 />
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">버튼모바일주소2</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="button_url_mobile2" value="<%= button_url_mobile2 %>" size="120" maxlength=256 />
                예) https://tenten.app.link/J3xFnMMFT4
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">실패시문자발송여부</td>
            <td bgcolor="#FFFFFF">
                <% Drawfailed_type "failed_type", failed_type, "" %>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">실패시문자제목</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="failed_subject" value="<%= failed_subject %>" size="55" maxlength=50 />
				<br>※ 입력시 실제 고객 데이터로 치환됨.
				<br>고객ID : <font color="red">${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}</font>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">실패시문자내용</td>
            <td bgcolor="#FFFFFF">
                <textarea name="failed_msg" cols=100 rows=8><%= failed_msg %></textarea>
				<br>※ 입력시 실제 고객 데이터로 치환됨.
				<br>고객ID : <font color="red">${CUSTOMERID},${CUSTOMERNAME},${CUSTOMERLEVELNAME}</font>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">사용여부</td>
            <td bgcolor="#FFFFFF">
                <% drawSelectBoxisusingYN "isusing", isusing, "" %>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">정렬순서</td>
            <td bgcolor="#FFFFFF">
                <input type="text" class="text" name="sortno" value="<%= sortno %>" size="5" maxlength=10 />
                예) 0
            </td>
        </tr>
        <% if tidx<>"" then %>
            <tr>
                <td bgcolor="#EFEFEF" align="center">수정로그</td>
                <td bgcolor="#FFFFFF">
                    최초작성 : <%= regadminid %>(<%= regdate %>)
                    <br>마지막수정 : <%= lastadminid %>(<%= lastupdate %>)
                </td>
            </tr>
        <% end if %>
        <tr>
            <td bgcolor="#FFFFFF" colspan=2 align="center">
                <input type="button" class="button" value="저장" onclick="jsRegCode();">
                &nbsp;
                <input type="button" class="button" value="신규등록" onclick="jsSetCode('');">
            </td>
        </tr>
        </table>		
        </form>
	</td>
</tr>
<tr>
	<td>
        <form name="frmSearch" method="post" action="/admin/appmanage/lms/lmstarget.asp" style="margin:0px;">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
        <tr bgcolor="#FFFFFF">
            <td colspan="8">
				발송방법 : <% Drawsendmethod "searchsendmethod", searchsendmethod, " onchange='jsSearch();'", "Y" %>
			</td>
        </tr>
		<tr bgcolor="#EFEFEF" align="center">
            <td>템플릿번호</td>
            <td>발송방법</td>
            <td>템플릿코드</td>
            <td>템플릿명</td>
			<td>내용</td>
            <td>정렬순서</td>
            <td>사용여부</td>
            <td>비고</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
            <tr bgcolor="#FFFFFF" align="center">
                <td><%=arrList(0,intLoop)%></td>
                <td><%=arrList(1,intLoop)%></td>
                <td><%=arrList(2,intLoop)%></td>
                <td align="left"><%=arrList(3,intLoop)%></td>
                <td align="left"><%= chrbyte(arrList(4,intLoop),50,"Y") %></td>
                <td align="left"><%=arrList(17,intLoop)%></td>
                <td><%=arrList(12,intLoop)%></td>
                <td>
                    <input type="button" class="button" value="수정" onclick="jsSetCode('<%=arrList(0,intLoop)%>');">
                </td>
            </tr>
			<%Next%>
		<%ELSE%>	
		<tr bgcolor="#FFFFFF">			
			<td colspan="8" align="center">등록된 내용이 없습니다.</td>
		</tr>	
		<%End if%>		
		</table>
        </form>
	</td>
</tr>
</table>

<%
session.codePage = 949
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->