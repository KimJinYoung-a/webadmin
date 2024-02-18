<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
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

function getYeoinSongjangDiv(songjangdiv)
    dim sdivcd 
    sdivcd = Trim(songjangdiv)
    
    Select case sdivcd
        case "24"   '' 사가와 익스프레스
            : getYeoinSongjangDiv = "D260"
        case "4"    '' CJ
            : getYeoinSongjangDiv = "D500"
        case "13"    '' 옐로우캡
            : getYeoinSongjangDiv = "D501"
        case "2"    '' 현대
            : getYeoinSongjangDiv = "D502"
        case "3"    '' 대한통운
            : getYeoinSongjangDiv = "D503"
        case "18"    '' 로젠택배
            : getYeoinSongjangDiv = "D504"
        case "7"    '' 훼미리넷
            : getYeoinSongjangDiv = "D505"
        case "9"    '' KGB택배
            : getYeoinSongjangDiv = "D506"
        case "10"    '' 아주
            : getYeoinSongjangDiv = "D507"
        case "20"    '' KT로지스
            : getYeoinSongjangDiv = "D508"
        case "5"    '' 이클라인
            : getYeoinSongjangDiv = "D509"
        case "21"    '' 경동택배
            : getYeoinSongjangDiv = "D510"
        case "17"   '' Tranet
            : getYeoinSongjangDiv = "D511" 
        case ".."   '' 양양택배
            : getYeoinSongjangDiv = "D512" 
        case "26"    ''일양택배
            : getYeoinSongjangDiv = "D513"
        case "8"    ''우체국택배
            : getYeoinSongjangDiv = "D514"
        case "1"    ''한진택배
            : getYeoinSongjangDiv = "D503"  ''수정요망
        case "6"    ''HTH
            : getYeoinSongjangDiv = "D503"  ''수정요망
        case "27"    ''LOEX택배
            : getYeoinSongjangDiv = "D503"  ''수정요망
        case "23"    ''신세계SEDEX
            : getYeoinSongjangDiv = "D503"  ''수정요망
        case "99"    ''신세계SEDEX
            : getYeoinSongjangDiv = "D503"  ''수정요망    
        case "25"    ''신세계SEDEX
            : getYeoinSongjangDiv = "D503"  ''수정요망    
        
        case ".."   ''건영택배
            : getYeoinSongjangDiv = "D515"
        case else
            : getYeoinSongjangDiv = ""
    end Select
    
''1	한진택배
''2	현대택배
''3	대한통운
''4	CJ GLS
''5	이클라인
''6	HTH
''7	훼미리택배
''8	우체국
''9	KGB택배
''10	아주택배
''11	오렌지택배
''12	한국택배
''13	옐로우캡
''14	나이스택배
''15	중앙택배
''16	주코택배
''17	트라넷택배
''18	로젠택배
''19	KGB특급택배
''20	KT로지스
''21	경동택배
''22	고려택배
''23	신세계 SEDEX
''24	사가와익스프레스
''25	하나로택배
''26	일양택배
''99	기타
end function

dim orderserial, extorderserial, mode
orderserial     = request("orderserial")
extorderserial  = request("extorderserial")
mode            = request("mode")

dim call_url, sedata
Const DELIMROW = "Y|R|T"
Const DELIMCOL = "Y|C|T"



''대표 송장 query
dim sqlStr
dim songjangNo, songjangDiv, yeoinSongjangDiv
dim delivername

sqlStr = "select count(d.idx) as cnt, d.songjangdiv, d.songjangno, s.divname, d.isupchebeasong"
sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_detail d"
sqlStr = sqlStr + " left join [db_order].[dbo].tbl_songjang_div s on d.songjangdiv=s.divcd"
sqlStr = sqlStr + " where orderserial='" + orderserial + "'"
sqlStr = sqlStr + " and itemid<>0"
sqlStr = sqlStr + " and cancelyn<>'Y'"
sqlStr = sqlStr + " and currstate='7'"
sqlStr = sqlStr + " group by d.songjangdiv, d.songjangno, s.divname, d.isupchebeasong"
sqlStr = sqlStr + " order by cnt desc, d.isupchebeasong "

rsget.Open sqlStr, dbget, 1
if Not rsget.Eof then
    songjangdiv = rsget("songjangdiv")
    songjangno  = rsget("songjangno")
    delivername = db2html(rsget("divname"))
    
    songjangno = replace(songjangno,"-","")
end if
rsget.Close

if (songjangdiv="") or (songjangno="") then
    response.write "<script >alert('배송정보가 부족합니다. songjangdiv : " + songjangdiv + ", songjangno : " + songjangno + "');</script>"
    dbget.close()	:	response.End 
end if

yeoinSongjangDiv = getYeoinSongjangDiv(songjangdiv)

if (yeoinSongjangDiv="") then
    response.write "<script >alert('택배사가 정의되지 않았습니다. songjangdiv : " + songjangdiv + "');</script>"
    dbget.close()	:	response.End 
end if

call_url = "http://www.yeoin.com/site/tenbyten/TenByTen_OrderStatus_.jsp"
sedata = "sParam=REGI_DELI" & DELIMCOL & extorderserial & DELIMCOL & songjangNo & DELIMCOL & yeoinSongjangDiv '' 명령어 Y|C|T 주문번호 Y|C|T 송장번호 Y|C|T 택배사코드 


dim reText
if (mode="senddata") then
    reText = Trim(SendReq(call_url,sedata))
    
    '' 결과 주문번호Y|C|T1  :1 정상, 0 실패
    if (InStr(reText,extorderserial & "Y|C|T1")>=0) then
        response.write "<script>opener.location.reload();</script>"
    else
        response.write "<script>alert('저장 오류.');</script>"
    end if
end if

%>
<script language='javascript'>
function SaveSongjang(frm){
    if (confirm('전송 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmSubmit" method="post" action="">
<input type="hidden" name="mode" value="senddata">
<tr>
    <td colspan="2" bgcolor="#FFFFFF"><%= call_url %></td>
</tr>
<tr>
    <td colspan="2" bgcolor="#FFFFFF"><%= sedata %></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">원주문번호</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="extorderserial" value="<%= extorderserial %>" size="26" readOnly ></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">텐바이텐<br>주문번호</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="26" readOnly ></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">송장번호</td>
    <td bgcolor="#FFFFFF"><%= songjangno %></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">택배사</td>
    <td bgcolor="#FFFFFF"><%= yeoinSongjangDiv %> (<%= delivername %>)</td>
</tr>
<% if (mode="senddata") then %>
<tr>
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <%= reText %>
    </td>
</tr>
<% else %>
<tr>
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <input type="button" class="button" value=" 송장 전송 " onclick="SaveSongjang(frmSubmit);" onFocus="this.blur();">
    </td>
</tr>
<% end if %>

</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->