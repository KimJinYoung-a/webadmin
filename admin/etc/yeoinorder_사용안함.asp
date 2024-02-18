<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->



<%
Function SendReq(call_url, sedata)
    dim objHttp, ret_txt
    Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")
    objHttp.Open "POST", call_url, False
    objHttp.setRequestHeader "Connection", "close"
    objHttp.setRequestHeader "Content-Length", Len(sedata)
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.Send  sedata
    ret_txt = objHttp.ResponseBody
    set objHttp = Nothing
    
    SendReq = Trim(BinToText(ret_txt,8192))
end function

Function BinToText(varBinData, intDataSizeBytes)
	Const adFldLong = &H00000080
	Const adVarChar = 200

	dim objRS, strV, tmpMsg,isError

	Set objRS = CreateObject("ADODB.Recordset")
	objRS.Fields.Append "txt", adVarChar, intDataSizeBytes, adFldLong
	objRS.Open
	objRS.AddNew
	objRS.Fields("txt").AppendChunk varBinData
	strV=objRS("txt").Value
	BinToText = strV
	objRS.Close
	Set objRS=Nothing
End Function

Function StripTags(htmlDoc)
	Dim rex
	Set rex = new Regexp
	rex.Pattern= "<[^>]+>"
	rex.Global=True
	StripTags =rex.Replace(htmlDoc,"")
	Set rex = Nothing
End Function


dim i, mode, research
mode        = request("mode")
research    = request("research")

''DefaultSetting 
if (research="") and (mode="") then mode="00002"


dim objHttp
dim bufStr
dim call_url, sedata, ret_txt


''조회 시작일 1달전부터.
dim sStartDay
sStartDay = Replace(Left(DateAdd("d",-61,now()),10),"-","")

call_url = "http://www.yeoin.com/site/tenbyten/TenByTen_OrderInfo_.jsp"
sedata   = "sStatusCd=" + mode + "&sStartDay=" + sStartDay + "&sEndDay="    '' 실제 운영시 00002 : 결제완료

''response.write call_url+"?" +sedata

if mode<>"" then
    ret_txt = SendReq(call_url, sedata)
    ret_txt = replace(ret_txt," " & VbCrlf , "")
    ret_txt = replace(ret_txt,VbCrlf , "")
    ret_txt = replace(ret_txt,"<!--------------------- 공통Bottom Start ----------------------->" , "")
    ret_txt = replace(ret_txt,"[가격 미표기]" , "")
end if

dim RowData, rowcount
Const DELIMROW = "Y|R|T"
Const DELIMCOL = "Y|C|T"

if (ret_txt<>"") then
    RowData = split(ret_txt,DELIMROW)
end if

if IsArray(rowdata) then
    rowcount = UBound(rowdata) + 1
else
    rowcount = 0
end if


dim ColumnValue
dim PreCkNum

''최초주문일 : 통합주문데이타와 비교하기위해 ..
dim LastRow, LASTOrderDate
dim ExistsExtOrderSerialArr
dim sqlStr


if rowcount>0 then
    LastRow = RowData(rowcount-1)
    LastRow = split(Trim(LastRow),DELIMCOL)
    LASTOrderDate = LastRow(5)
    LASTOrderDate = Left(LASTOrderDate,4) + "-" + Mid(LASTOrderDate,5,2) + "-" + Mid(LASTOrderDate,7,2)
    
    response.write "LASTOrderDate : " & LASTOrderDate
    ''''테스트 후 삭제
    ''LASTOrderDate = "2007-01-01"
    ''''
    
    sqlStr = " select distinct m.orderserial, m.authcode, m.ipkumdiv, m.deliverno " + VbCrlf
    sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m" + VbCrlf
    sqlStr = sqlStr + " where m.regdate>='" + LASTOrderDate + "'" + VbCrlf
    sqlStr = sqlStr + " and m.sitename='yeoin'"
    sqlStr = sqlStr + " and m.cancelyn='N'"
    
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        ExistsExtOrderSerialArr = rsget.getRows()
    end if
    rsget.Close
end if

'If IsArray(ExistsExtOrderSerialArr) Then
'    'response.write Ubound(ExistsExtOrderSerialArr,2)
'    For i = 0 To Ubound(ExistsExtOrderSerialArr,2)
'        response.write ExistsExtOrderSerialArr(0,i) & "<br>"
'    Next
'End If



''tbl_order_master.authcode 에 제휴사 주문번호 저장


dim IsAlreadySaved
dim tenorderserial, tenipkumdiv, tendeliveryno

function fnIsAleadySenedData(iExistsExtOrderSerialArr, extorderserial, byref tenorderserial, byref tenipkumdiv, byref tendeliveryno)
    dim cnt, i
    
    If IsArray(iExistsExtOrderSerialArr) Then
        cnt = Ubound(iExistsExtOrderSerialArr,2)
        
        For i = 0 To cnt
            if Trim(iExistsExtOrderSerialArr(1,i))=Trim(extorderserial) then
                fnIsAleadySenedData = true
                tenorderserial  = iExistsExtOrderSerialArr(0,i)
                tenipkumdiv     = iExistsExtOrderSerialArr(2,i)
                tendeliveryno   = iExistsExtOrderSerialArr(3,i)
                Exit function
            end if
        Next
    end if
    fnIsAleadySenedData = false
end function


%>
<script languaga='javascript'>
function reSearch(frm){
    frm.submit();
}

function SaveExtOrder(extorderserial){
    var frm = document.frmSubmit;
    
    if (confirm('텐바이텐 주문 목록에 저장 하시겠습니까?')){
        frm.action = "yeoinorder_process.asp";
        frm.extorderserial.value = extorderserial;
        frm.submit();
    }
}

function SendSongjang(extorderserial, orderserial){
    var popwin = window.open('popyeoinorder_songjanginput.asp?orderserial=' + orderserial + '&extorderserial=' + extorderserial,'SendSongjang','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="GET" action="" onSubmit="return false">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" >
	    <select name="mode">
	    <option value="" >선택
	    <option value="00001" <% if mode="00001" then response.write "selected" %> >처리대기중 
	    <option value="00002" <% if mode="00002" then response.write "selected" %> >결제완료
	    <option value="00003" <% if mode="00003" then response.write "selected" %> >상품포장중
	    <option value="00004" <% if mode="00004" then response.write "selected" %> >출고완료
	    <option value="00005" <% if mode="00005" then response.write "selected" %> >처리대기중[취소]
	    <option value="00006" <% if mode="00006" then response.write "selected" %> >결제완료[취소] 
	    <option value="00007" <% if mode="00007" then response.write "selected" %> >상품포장중[취소]
	    <option value="00008" <% if mode="00008" then response.write "selected" %> >출고완료[취소] 
	    </select>
       	<img src="/admin/images/search2.gif" onClick="reSearch(frm_search)" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 상단 검색폼 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmSubmit" method="post" action="">
<input type="hidden" name="extorderserial" value="">
<tr bgcolor="#DDDDFF">
    <td colspan="24">여인닷컴 DATA <%= call_url %>?<%= sedata %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="24"><textarea name="ORGData" rows="4" cols="120" readonly ><%= ret_txt %></textarea></td>
</tr>
<tr bgcolor="#DDDDFF">
    <td colspan="24">Parsing DATA</td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
<!--
	<td align="center">주문상태명</td>
	<td align="center">결제방식</td>
	<td align="center">주문상태코드</td>
-->	
	<td align="center">주문통합코드<br>주문코드<br>주문일시</td>
	<td align="center" width="40">상품<br>코드</td>
<!--	
	<td align="center">여인닷컴 상품코드</td>
-->		
	<td align="center" >상품명</td>
	<td align="center" >옵션</td>
	<td align="center">수량</td>
	<td align="center">단가</td>
	<!-- <td align="center">주문금액</td> -->
	<td align="center">주문자<br>수령인</td>
	<td align="center">전화번호</td>
	<td align="center">휴대번호</td>
	<td align="center">주소</td>
	<td align="center">배송메세지</td>
	<td align="center">배송비</td>
	<td align="center">송장번호</td>
	<td align="center">배송업체 코드</td>
	<td align="center">배송업체 명</td>
	<td align="center">주문저장</td>
	<td align="center">송장저장</td>
</tr>

<% for i=0 to rowCount - 1 %>
<% 
    ColumnValue = split(Trim(RowData(i)),DELIMCOL)
    if IsArray(ColumnValue) then
%>

<% if (PreCkNum<>"") and (PreCkNum<>ColumnValue(4)) then %>
<tr>
    <td colspan="24" bgcolor="#DD33FF" height="1"></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
<!--
    <td><%= ColumnValue(0) %></td>
    <td><%= ColumnValue(1) %></td>
    <td><%= ColumnValue(2) %></td>
-->
    <td><%= ColumnValue(3) %><br><%= ColumnValue(4) %><br><%= ColumnValue(5) %></td>
    <td><%= ColumnValue(6) %></td> <!-- 텐바이텐 상품코드 -->
    
    <!-- <td><%= ColumnValue(7) %></td> -->  
    <td><%= ColumnValue(8) %></td>  <!-- 상품명 -->
    <td><%= ColumnValue(9) %></td>  <!-- 옵션 -->
    <td align="center"><%= ColumnValue(10) %></td>  <!-- 수량 -->
    <td align="right"><%= ColumnValue(11) %></td>   <!-- 단가 -->
    <!-- <td align="right"><%= ColumnValue(12) %></td>   -->
    <td><%= ColumnValue(23) %><br><%= ColumnValue(13) %></td>     <!-- 주문자 / 수령인 -->
    <td><%= ColumnValue(24) %><br><%= ColumnValue(14) %></td>     <!-- 주문자 전화 / 수령인 전화 -->
    <td><%= ColumnValue(25) %><br><%= ColumnValue(15) %></td>     <!-- 주문자 핸드폰 / 수령인 핸드폰 -->
    <td>
        <%= ColumnValue(16) %>          <!-- 우편번호 -->
        <br>
        <%= ColumnValue(17) %>          <!-- 주소 -->
    </td>
    <td><textarea class="textarea" name="msg" cols="10" rows="3"><%= ColumnValue(18) %></textarea></td> <!-- 배송메세지 -->
    <td><%= ColumnValue(19) %></td>     <!-- 배송비 -->
    <td><%= ColumnValue(20) %></td>
    <td><%= ColumnValue(21) %></td>
    <td><%= ColumnValue(22) %></td>
    <% if (PreCkNum<>ColumnValue(4)) then %>
    <%
        if (mode="00002") then
            IsAlreadySaved = fnIsAleadySenedData(ExistsExtOrderSerialArr,ColumnValue(4), tenorderserial, tenipkumdiv, tendeliveryno)
        else
            IsAlreadySaved = false
            tenorderserial  = ""
            tenipkumdiv     = ""
            tendeliveryno   = ""
        end if
    %>
    <td>
        <% if (mode="00002") and (Not IsAlreadySaved) then %>
        <input class="button" type="button" value="저장" onClick="SaveExtOrder('<%= ColumnValue(4) %>')" onFocus="this.blur()">
        <% else %>
        <a href="/admin/ordermaster/viewordermaster.asp?orderserial=<%= tenorderserial %>" target="_blank"><%= tenorderserial %></a>
        <% end if %>
    </td>
    <td>
        <% if (mode="00002") and (IsAlreadySaved) and (tenipkumdiv="8") then  %>
        <input class="button" type="button" value="전송" onClick="SendSongjang('<%= ColumnValue(4) %>','<%= tenorderserial %>')" onFocus="this.blur()">
        <% end if %>
    </td>
    <% else %>
    <td ></td>
    <td ></td>
    <% end if %>
</tr>
    <% 
        PreCkNum = ColumnValue(4)
    end if
    %>
<% next %>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->