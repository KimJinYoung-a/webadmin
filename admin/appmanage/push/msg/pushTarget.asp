<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 타켓 관리
' Hieditor : 2019.06.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->

<%
Dim arrList,intLoop, clsCode, sMode, searchrepeatpushyn, replacetagcode
dim targetKey,targetName,targetQuery,isusing,repeatpushyn
    targetkey = requestcheckvar(request("targetkey"),10)
    searchrepeatpushyn = requestcheckvar(request("searchrepeatpushyn"),1)
	sMode ="I"
	
Set clsCode = new CpushtargetCommonCode  	
	IF targetkey <> "" THEN
		sMode ="U"
		clsCode.frecttargetKey  = targetKey 
		clsCode.GetpushtargetCont

        if clsCode.FTotalCount>0 THEN
            targetName = clsCode.ftargetName
            targetQuery = clsCode.ftargetQuery
            isusing = clsCode.fisusing
            repeatpushyn = clsCode.frepeatpushyn
			replacetagcode = clsCode.freplacetagcode
        end if
    END IF
 		 
	clsCode.frectrepeatpushyn = searchrepeatpushyn
	arrList = clsCode.GetpushtargetList
Set clsCode = nothing 
 
%>
<script type='text/javascript'>

	// 코드타입 변경이동
	function jsSetCode(targetkey, searchrepeatpushyn){	
		self.location.href = "/admin/appmanage/push/msg/pushtarget.asp?targetkey="+targetkey+"&searchrepeatpushyn="+searchrepeatpushyn;
	}
	
	//코드 검색
	function jsSearch(){
		document.frmSearch.submit();
	}
	
	//코드 등록
	function jsRegCode(){
		var frm = document.frmReg;
		if(!frm.targetKey.value) {
			alert("타켓키를 입력해 주세요");
			frm.targetKey.focus();
			return false;
		}
		if(!frm.targetName.value) {
			alert("타켓이름을 입력해 주세요");
			frm.targetName.focus();
			return false;
		}
		if(!frm.isusing.value) {
			alert("사용여부를 입력해 주세요");
			frm.isusing.focus();
			return false;
		}
		if(!frm.repeatpushyn.value) {
			alert("푸시구분을 입력해 주세요");
			frm.repeatpushyn.focus();
			return false;
		}
		frm.submit();
	}
	
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a" >
<tr>
	<td>
		<form name="frmReg" method="post" action="/admin/appmanage/push/msg/pushtarget_process.asp" style="margin:0px;">	
		<input type="hidden" name="sM" value="<%=sMode%>">
        <table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">										
        <tr>
            <td bgcolor="#EFEFEF" width="100" align="center">타켓키</td>
            <td bgcolor="#FFFFFF">
                <% if targetKey<>"" then %>
                    <%= targetkey %>
                    <input type="hidden" size="8" maxlength="10" name="targetKey" value="<%= targetkey %>">
                <% else %>
                    <input type="text" size="8" maxlength="10" name="targetKey" value="<%= targetkey %>">
                <% end if %>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">타켓이름</td>
            <td bgcolor="#FFFFFF">
                <input type="text" size="100" maxlength="100" name="targetName" value="<%= targetName %>">
            </td>
        </tr>					
        <tr>
            <td bgcolor="#EFEFEF" align="center">쿼리</td>
            <td bgcolor="#FFFFFF">
                <textarea name="targetQuery" cols=100 rows=8><%= targetQuery %></textarea>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">치환코드</td>
            <td bgcolor="#FFFFFF">
                <textarea name="replacetagcode" cols=100 rows=2><%= replacetagcode %></textarea>
				<br>예) ${CUSTOMERID}
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">사용여부</td>
            <td bgcolor="#FFFFFF">
                <% drawSelectBoxisusingYN "isusing", isusing, "" %>
            </td>
        </tr>
        <tr>
            <td bgcolor="#EFEFEF" align="center">푸시구분</td>
            <td bgcolor="#FFFFFF">
				<% Drawpushgubun "repeatpushyn", repeatpushyn, "", "" %>
            </td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF" colspan=2 align="center">
                <input type="button" class="button" value="저장" onclick="jsRegCode();">
                &nbsp;
                <input type="button" class="button" value="신규등록" onclick="jsSetCode('','<%= searchrepeatpushyn %>');">
            </td>
        </tr>
        </table>		
        </form>
	</td>
</tr>
<tr>
	<td>
        <form name="frmSearch" method="post" action="/admin/appmanage/push/msg/pushtarget.asp" style="margin:0px;">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">				
        <tr bgcolor="#FFFFFF">
            <td colspan="7">
				<% Drawpushgubun "searchrepeatpushyn", searchrepeatpushyn, " onchange='jsSearch();'", "Y" %>
			</td>
        </tr>
		<tr bgcolor="#EFEFEF">
			<td align="center" width="60">타켓키</td>
			<td align="center" width="200">타켓이름</td>
			<td align="center">쿼리</td>
			<td align="center">치환코드</td>
			<td align="center" width="40">사용<br>여부</td>
			<td align="center" width="40">반복<br>푸시<br>여부</td>
            <td align="center" width="40">비고</td>
		</tr>
		<%If isArray(arrList) THEN%>
			<%For intLoop = 0 To UBound(arrList,2)%>
		<tr bgcolor="#FFFFFF">
			<td align="center"><%=arrList(0,intLoop)%></td>
			<td align="left"><%=arrList(1,intLoop)%></td>
			<td align="left"><%=arrList(2,intLoop)%></td>
			<td align="left"><%=arrList(6,intLoop)%></td>
			<td align="center"><%=arrList(3,intLoop)%></td>
            <td align="center"><%=arrList(4,intLoop)%></td>
			<td align="center">
                <input type="button" class="button" value="수정" onclick="jsSetCode('<%=arrList(0,intLoop)%>','<%= searchrepeatpushyn %>');">
			</td>
		</tr>
			<%Next%>
		<%ELSE%>	
		<tr bgcolor="#FFFFFF">			
			<td colspan="5" align="center">등록된 내용이 없습니다.</td>
		</tr>	
		<%End if%>		
		</table>
        </form>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->