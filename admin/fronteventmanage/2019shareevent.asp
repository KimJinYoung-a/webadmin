<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 4월 정기세일 - 앗싸! 에어팟2 득템 응모자페이지
' History : 2019-04-11 이종화
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim sqlStr , mktArr , intLoop

sqlStr = "EXEC db_temp.dbo.usp_WWW_snsshare_mktdata"
rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly	
if not rsget.EOF then
    mktArr = rsget.getRows()	
end if
rsget.close
%>
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<link rel="stylesheet" type="text/css" href="/js/jqueryui/css/jquery-ui.css"/>
<div class="content scrl" style="top:40px;">
	<div class="pad20">
        <table class="tbType1 listTb" style="width:800px;margin-left:50;">
            <colgroup>
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
                <col width="16%" />
            </colgroup>
            <tr align="center" bgcolor="#E6E6E6" height="20">
                <th><strong>날짜</strong></th>
                <th><strong>시간</strong></th>
                <th><strong>이미지 받은 유저</strong></th>
                <th><strong>이미지 받은 횟수</strong></th>
                <th><strong>응모자수</strong></th>
                <th><strong>응모횟수</strong></th>
            </tr>
            <% IF isArray(mktArr) THEN %>
            <% For intLoop = 0 To UBound(mktArr,2) %>
            <tr>
                <td><%=mktArr(0,intLoop)%></td>
                <td><%=mktArr(1,intLoop)%></td>
                <td><%=mktArr(2,intLoop)%></td>
                <td><%=mktArr(3,intLoop)%></td>
                <td><%=mktArr(4,intLoop)%></td>
                <td><%=mktArr(5,intLoop)%></td>
            </tr>
            <% Next %>
            <% End If %>
        </table>
    </div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->