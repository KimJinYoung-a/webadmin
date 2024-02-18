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
        case "24"   '' �簡�� �ͽ�������
            : getYeoinSongjangDiv = "D260"
        case "4"    '' CJ
            : getYeoinSongjangDiv = "D500"
        case "13"    '' ���ο�ĸ
            : getYeoinSongjangDiv = "D501"
        case "2"    '' ����
            : getYeoinSongjangDiv = "D502"
        case "3"    '' �������
            : getYeoinSongjangDiv = "D503"
        case "18"    '' �����ù�
            : getYeoinSongjangDiv = "D504"
        case "7"    '' �ѹ̸���
            : getYeoinSongjangDiv = "D505"
        case "9"    '' KGB�ù�
            : getYeoinSongjangDiv = "D506"
        case "10"    '' ����
            : getYeoinSongjangDiv = "D507"
        case "20"    '' KT������
            : getYeoinSongjangDiv = "D508"
        case "5"    '' ��Ŭ����
            : getYeoinSongjangDiv = "D509"
        case "21"    '' �浿�ù�
            : getYeoinSongjangDiv = "D510"
        case "17"   '' Tranet
            : getYeoinSongjangDiv = "D511" 
        case ".."   '' ����ù�
            : getYeoinSongjangDiv = "D512" 
        case "26"    ''�Ͼ��ù�
            : getYeoinSongjangDiv = "D513"
        case "8"    ''��ü���ù�
            : getYeoinSongjangDiv = "D514"
        case "1"    ''�����ù�
            : getYeoinSongjangDiv = "D503"  ''�������
        case "6"    ''HTH
            : getYeoinSongjangDiv = "D503"  ''�������
        case "27"    ''LOEX�ù�
            : getYeoinSongjangDiv = "D503"  ''�������
        case "23"    ''�ż���SEDEX
            : getYeoinSongjangDiv = "D503"  ''�������
        case "99"    ''�ż���SEDEX
            : getYeoinSongjangDiv = "D503"  ''�������    
        case "25"    ''�ż���SEDEX
            : getYeoinSongjangDiv = "D503"  ''�������    
        
        case ".."   ''�ǿ��ù�
            : getYeoinSongjangDiv = "D515"
        case else
            : getYeoinSongjangDiv = ""
    end Select
    
''1	�����ù�
''2	�����ù�
''3	�������
''4	CJ GLS
''5	��Ŭ����
''6	HTH
''7	�ѹ̸��ù�
''8	��ü��
''9	KGB�ù�
''10	�����ù�
''11	�������ù�
''12	�ѱ��ù�
''13	���ο�ĸ
''14	���̽��ù�
''15	�߾��ù�
''16	�����ù�
''17	Ʈ����ù�
''18	�����ù�
''19	KGBƯ���ù�
''20	KT������
''21	�浿�ù�
''22	����ù�
''23	�ż��� SEDEX
''24	�簡���ͽ�������
''25	�ϳ����ù�
''26	�Ͼ��ù�
''99	��Ÿ
end function

dim orderserial, extorderserial, mode
orderserial     = request("orderserial")
extorderserial  = request("extorderserial")
mode            = request("mode")

dim call_url, sedata
Const DELIMROW = "Y|R|T"
Const DELIMCOL = "Y|C|T"



''��ǥ ���� query
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
    response.write "<script >alert('��������� �����մϴ�. songjangdiv : " + songjangdiv + ", songjangno : " + songjangno + "');</script>"
    dbget.close()	:	response.End 
end if

yeoinSongjangDiv = getYeoinSongjangDiv(songjangdiv)

if (yeoinSongjangDiv="") then
    response.write "<script >alert('�ù�簡 ���ǵ��� �ʾҽ��ϴ�. songjangdiv : " + songjangdiv + "');</script>"
    dbget.close()	:	response.End 
end if

call_url = "http://www.yeoin.com/site/tenbyten/TenByTen_OrderStatus_.jsp"
sedata = "sParam=REGI_DELI" & DELIMCOL & extorderserial & DELIMCOL & songjangNo & DELIMCOL & yeoinSongjangDiv '' ��ɾ� Y|C|T �ֹ���ȣ Y|C|T �����ȣ Y|C|T �ù���ڵ� 


dim reText
if (mode="senddata") then
    reText = Trim(SendReq(call_url,sedata))
    
    '' ��� �ֹ���ȣY|C|T1  :1 ����, 0 ����
    if (InStr(reText,extorderserial & "Y|C|T1")>=0) then
        response.write "<script>opener.location.reload();</script>"
    else
        response.write "<script>alert('���� ����.');</script>"
    end if
end if

%>
<script language='javascript'>
function SaveSongjang(frm){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
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
    <td width="100" bgcolor="#DDDDFF">���ֹ���ȣ</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="extorderserial" value="<%= extorderserial %>" size="26" readOnly ></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">�ٹ�����<br>�ֹ���ȣ</td>
    <td bgcolor="#FFFFFF"><input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="26" readOnly ></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">�����ȣ</td>
    <td bgcolor="#FFFFFF"><%= songjangno %></td>
</tr>
<tr>
    <td width="100" bgcolor="#DDDDFF">�ù��</td>
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
    <input type="button" class="button" value=" ���� ���� " onclick="SaveSongjang(frmSubmit);" onFocus="this.blur();">
    </td>
</tr>
<% end if %>

</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->