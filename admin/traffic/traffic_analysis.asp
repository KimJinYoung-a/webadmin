<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  다음추출  traffic analysis 페이지
' History : 2007.09.04 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->

<%		
dim seach,seach2 ,ColumnValue ,i 
	seach = request("seach")
	seach2 = request("seach2")
															
'################################################################################### 전체트래픽 추이시작
dim objHttp
Set objHttp = CreateObject("Msxml2.ServerXMLHTTP")

Function SendReq(call_url, sedata)
    dim ret_txt
    
    objHttp.Open "POST", call_url, False
    objHttp.setRequestHeader "Connection", "close"
    objHttp.setRequestHeader "Content-Length", Len(sedata)
    objHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHttp.Send  sedata
    ret_txt = objHttp.ResponseBody
    
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

dim bufStr
dim call_url, sedata, ret_txt
''call_url = "http://login.daum.net/Mail-bin/login.cgi?id=tozzinet&pw=g3d6a9&daumauth=1&service=&category=webinside"	'구로그인
call_url = "https://logins.daum.net/accounts/login.do?id=dostardom&pw=1dzsback&url=http://inside.daum.net/dwi/report/top/Summary.dwi&webmsg=-1"		'2010로그인(박도)
ret_txt = SendReq(call_url, sedata)

call_url = "http://inside.daum.net/dwi/report/traffic/All.dwi?mode=0&fromDate="+seach2+"&reportType=1&toDate="+seach
ret_txt = SendReq(call_url, sedata)

dim RowData, rowcount

if (ret_txt<>"") then
	 ret_txt  = replace(ret_txt,"chr(13) & chr(10)"," ")	'10개
	 ret_txt  = replace(ret_txt,vbTab," ")
	 ret_txt  = replace(ret_txt,"          "," ")	'10개
	 ret_txt  = replace(ret_txt,"         "," ")	'9개
	 ret_txt  = replace(ret_txt,"        "," ")		'8개
	 ret_txt  = replace(ret_txt,"       "," ")		'7개
	 ret_txt  = replace(ret_txt,"      "," ")		'6개
	 ret_txt  = replace(ret_txt,"     "," ")		'5개
	 ret_txt  = replace(ret_txt,"    "," ")			'4개
	 ret_txt  = replace(ret_txt,"   "," ")			'3개
	 ret_txt  = replace(ret_txt,"  "," ")			'2개 
	 
    RowData = split(ret_txt," ") 
end if

if IsArray(rowdata) then
    rowcount = UBound(rowdata) + 1
else
    rowcount = 0
end if

'################################################################################### 실제방문자 추이시작
dim bufStr1 ,call_url1, sedata1, ret_txt1 ,RowData1, rowcount1 ,ColumnValue1
	call_url1 = "http://inside.daum.net/dwi/report/traffic/Visit.dwi?mode=0&fromDate="+seach2+"&reportType=1&toDate="+seach
    ret_txt1 = SendReq(call_url1, sedata1)

if (ret_txt1<>"") then
	 ret_txt1  = replace(ret_txt1,"chr(13) & chr(10)"," ")	'10개
	 ret_txt1  = replace(ret_txt1,vbTab," ")
	 ret_txt1  = replace(ret_txt1,"          "," ")	'10개
	 ret_txt1  = replace(ret_txt1,"         "," ")	'9개
	 ret_txt1  = replace(ret_txt1,"        "," ")		'8개
	 ret_txt1  = replace(ret_txt1,"       "," ")		'7개
	 ret_txt1  = replace(ret_txt1,"      "," ")		'6개
	 ret_txt1  = replace(ret_txt1,"     "," ")		'5개
	 ret_txt1  = replace(ret_txt1,"    "," ")			'4개
	 ret_txt1  = replace(ret_txt1,"   "," ")			'3개
	 ret_txt1  = replace(ret_txt1,"  "," ")			'2개 
	 
    RowData1 = split(ret_txt1," ") 
end if

if IsArray(RowData1) then
    rowcount1 = UBound(RowData1) + 1
else
    rowcount1 = 0
end if
 
dim call_url2, sedata2, ret_txt2 ,bufStr2 ,RowData2, rowcount2 ,ColumnValue2
	call_url2 = "http://inside.daum.net/dwi/report/traffic/Uv.dwi?mode=0&fromDate="+seach2+"&reportType=1&toDate="+seach
    ret_txt2 = SendReq(call_url2, sedata2)

if (ret_txt2<>"") then
	 ret_txt2  = replace(ret_txt2,"chr(13) & chr(10)"," ")	'10개
	 ret_txt2  = replace(ret_txt2,vbTab," ")
	 ret_txt2  = replace(ret_txt2,"          "," ")	'10개
	 ret_txt2  = replace(ret_txt2,"         "," ")	'9개
	 ret_txt2  = replace(ret_txt2,"        "," ")		'8개
	 ret_txt2  = replace(ret_txt2,"       "," ")		'7개
	 ret_txt2  = replace(ret_txt2,"      "," ")		'6개
	 ret_txt2  = replace(ret_txt2,"     "," ")		'5개
	 ret_txt2  = replace(ret_txt2,"    "," ")			'4개
	 ret_txt2  = replace(ret_txt2,"   "," ")			'3개
	 ret_txt2  = replace(ret_txt2,"  "," ")			'2개 
	 
    RowData2 = split(ret_txt2," ") 
end if

if IsArray(RowData2) then
    rowcount2 = UBound(RowData2) + 1
else
    rowcount2 = 0
end if
%> 

<script language="javascript">

//텐바이텐 db에 저장
function autosubmit()
{
	document.frm.action = "traffic_tenbyten_submit.asp";
	document.frm.submit();
}


//수동입력 시작
function sudongsubmit()
{
	document.frm.action = "traffic_analysis_sudong.asp";
	document.frm.submit();
}

//날짜 검색 조건
function seachform(seach2,seach2){

	if (!IsDouble(frm.seach2.value)){
		alert('시작일을 7일 단위로 정확히 입력하세요. 숫자만 가능합니다.');
		frm.seach2.focus();
		return;
	}
	if (!IsDouble(frm.seach.value)){
		alert('마지막일을 7일 단위로 정확히 입력하세요. 숫자만 가능합니다.');
		frm.seach.focus();
		return;
	}
	document.frm.submit();

}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="post">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		날짜 : <input type="text" name="seach2" size="8" value="<%= seach2 %>" maxlength=8> ~ 
		<input type="text" name="seach" size="8" value="<%= seach %>" maxlength=8> 		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="seachform('<%= seach2 %>','<%= seach %>');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">		
	</td>
</tr>
</table>
<!-- 검색 끝 -->
		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		ex) 20070101 ~ 20070107 <font color="red"><strong>7일단위로 정확이 입력 하셔야 검색이 됩니다.</strong></font>
	</td>
	<td align="right">
		<input type="button" value="수동입력" onclick="sudongsubmit()" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
if seach2 <> "" then

ColumnValue = RowData
ColumnValue1 = RowData1
ColumnValue2 = RowData2 	
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   <td >날짜</td>
   <td >페이지뷰</td>
   <td>방문자수</td>
   <td >신규방문자수</td>
   <td >재방문자수</td>
   <td >실제방문자수</td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td><%= ColumnValue(12) %><input type="hidden" name="ColumnValue_12" value="<%= ColumnValue(12) %>"></td>
	<td><%= ColumnValue(13) %><input type="hidden" name="ColumnValue_13" value="<%= ColumnValue(13) %>"></td>
	<td><%= ColumnValue(15) %><input type="hidden" name="ColumnValue_15" value="<%= ColumnValue(15) %>"></td>
	<td><%= ColumnValue1(10) %><input type="hidden" name="ColumnValue1_10" value="<%= ColumnValue1(10) %>"></td>
	<td><%= ColumnValue1(11) %><input type="hidden" name="ColumnValue1_11" value="<%= ColumnValue1(11) %>"></td>
	<td><%= ColumnValue2(9) %><input type="hidden" name="ColumnValue2_9" value="<%= ColumnValue2(9) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(17) %><input type="hidden" name="ColumnValue_17" value="<%= ColumnValue(17) %>"></td>
	<td ><%= ColumnValue(18) %><input type="hidden" name="ColumnValue_18" value="<%= ColumnValue(18) %>"></td>
	<td ><%= ColumnValue(20) %><input type="hidden" name="ColumnValue_20" value="<%= ColumnValue(20) %>"></td>
	<td ><%= ColumnValue1(14) %><input type="hidden" name="ColumnValue1_14" value="<%= ColumnValue1(14) %>"></td>
	<td ><%= ColumnValue1(15) %><input type="hidden" name="ColumnValue1_15" value="<%= ColumnValue1(15) %>"></td>
	<td ><%= ColumnValue2(12) %><input type="hidden" name="ColumnValue2_12" value="<%= ColumnValue2(12) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(22) %><input type="hidden" name="ColumnValue_22" value="<%= ColumnValue(22) %>"></td>
	<td ><%= ColumnValue(23) %><input type="hidden" name="ColumnValue_23" value="<%= ColumnValue(23) %>"></td>
	<td ><%= ColumnValue(25) %><input type="hidden" name="ColumnValue_25" value="<%= ColumnValue(25) %>"></td>
	<td ><%= ColumnValue1(18) %><input type="hidden" name="ColumnValue1_18" value="<%= ColumnValue1(18) %>"></td>
	<td ><%= ColumnValue1(19) %><input type="hidden" name="ColumnValue1_19" value="<%= ColumnValue1(19) %>"></td>
	<td ><%= ColumnValue2(15) %><input type="hidden" name="ColumnValue2_15" value="<%= ColumnValue2(15) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(27) %><input type="hidden" name="ColumnValue_27" value="<%= ColumnValue(27) %>"></td>
	<td ><%= ColumnValue(28) %><input type="hidden" name="ColumnValue_28" value="<%= ColumnValue(28) %>"></td>
	<td ><%= ColumnValue(30) %><input type="hidden" name="ColumnValue_30" value="<%= ColumnValue(30) %>"></td>
	<td ><%= ColumnValue1(22) %><input type="hidden" name="ColumnValue1_22" value="<%= ColumnValue1(22) %>"></td>
	<td ><%= ColumnValue1(23) %><input type="hidden" name="ColumnValue1_23" value="<%= ColumnValue1(23) %>"></td>
	<td ><%= ColumnValue2(18) %><input type="hidden" name="ColumnValue2_18" value="<%= ColumnValue2(18) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(32) %><input type="hidden" name="ColumnValue_32" value="<%= ColumnValue(32) %>"></td>
	<td ><%= ColumnValue(33) %><input type="hidden" name="ColumnValue_33" value="<%= ColumnValue(33) %>"></td>
	<td ><%= ColumnValue(35) %><input type="hidden" name="ColumnValue_35" value="<%= ColumnValue(35) %>"></td>
	<td ><%= ColumnValue1(26) %><input type="hidden" name="ColumnValue1_26" value="<%= ColumnValue1(26) %>"></td>
	<td ><%= ColumnValue1(27) %><input type="hidden" name="ColumnValue1_27" value="<%= ColumnValue1(27) %>"></td>
	<td ><%= ColumnValue2(21) %><input type="hidden" name="ColumnValue2_21" value="<%= ColumnValue2(21) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(37) %><input type="hidden" name="ColumnValue_37" value="<%= ColumnValue(37) %>"></td>
	<td ><%= ColumnValue(38) %><input type="hidden" name="ColumnValue_38" value="<%= ColumnValue(38) %>"></td>
	<td ><%= ColumnValue(40) %><input type="hidden" name="ColumnValue_40" value="<%= ColumnValue(40) %>"></td>
	<td ><%= ColumnValue1(30) %><input type="hidden" name="ColumnValue1_30" value="<%= ColumnValue1(30) %>"></td>
	<td ><%= ColumnValue1(31) %><input type="hidden" name="ColumnValue1_31" value="<%= ColumnValue1(31) %>"></td>
	<td ><%= ColumnValue2(24) %><input type="hidden" name="ColumnValue2_24" value="<%= ColumnValue2(24) %>"></td>
</tr>
<tr bgcolor="ffffff" align="center">
	<td ><%= ColumnValue(42) %><input type="hidden" name="ColumnValue_42" value="<%= ColumnValue(42) %>"></td>
	<td ><%= ColumnValue(43) %><input type="hidden" name="ColumnValue_43" value="<%= ColumnValue(43) %>"></td>
	<td ><%= ColumnValue(45) %><input type="hidden" name="ColumnValue_45" value="<%= ColumnValue(45) %>"></td>
	<td ><%= ColumnValue1(34) %><input type="hidden" name="ColumnValue1_34" value="<%= ColumnValue1(34) %>"></td>
	<td ><%= ColumnValue1(35) %><input type="hidden" name="ColumnValue1_35" value="<%= ColumnValue1(35) %>"></td>
	<td ><%= ColumnValue2(27) %><input type="hidden" name="ColumnValue2_27" value="<%= ColumnValue2(27) %>"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6><textarea name="ORGData" rows="11" cols="100" readonly ><%= ret_txt %></textarea> ※ 전체트래픽 추이</td>	
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6><textarea name="ORGData1" rows="11" cols="100" readonly ><%= ret_txt1 %></textarea> ※ 신규/재방문자 추이</td>	
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6><textarea name="ORGData2" rows="11" cols="100" readonly ><%= ret_txt2 %></textarea> ※ 실제방문자 추이</td>	
</tr>
<tr bgcolor=#FFFFFF>
   <td align="right" colspan=6><input type="button" value="텐바이텐 DB에 저장" onclick="autosubmit()" class="button"></td>
</tr>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
</tr>
<% end if %>
</form>
</table>

<% set objHttp = Nothing %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->