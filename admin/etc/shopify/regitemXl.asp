<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
public function getShopifyRows(byref colRows, byVal imakerid,byVal itemidArrs)
    Dim strSql 
    itemidArrs = replace(itemidArrs,vbcrlf,",")
    itemidArrs = replace(itemidArrs,vbcr,",")
    itemidArrs = replace(itemidArrs,vblf,",")
    rw itemidArrs
    strSql = " exec db_etcmall.[dbo].[usp_TEN_Shopify_regItemXL] '"&imakerid&"','"&itemidArrs&"'"
  
    rsget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.EOF Then
	    
	    colRows = Array()
	    For Each fld In rsget.Fields
	        reDim Preserve colRows(UBound(colRows) + 1)
            colRows(UBound(colRows))=fld.Name
            
        Next
        
		getShopifyRows = rsget.getRows()
	End If
	rsget.Close
	
end function

Dim vArrData, vArrCols, i, j
Dim itemids
Dim vSDate, vEDate, vChannel, vOrdType
Dim vDategbn, addparam1, addparam2
Dim makerid

makerid = requestCheckvar(request("makerid"),32)
itemids = requestCheckvar(request("itemids"),4000)
vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vDategbn = requestCheckvar(request("dategbn"),10)
vChannel = requestCheckvar(request("channel"),10)
vOrdType = requestCheckvar(request("ordtype"),32)
addparam1= requestCheckvar(request("addparam1"),32)
addparam2= requestCheckvar(request("cpngbn"),32)

if (vOrdType="") then vOrdType="S" ''건수(C) , 금액(S), 수익(G)

'' 기본값 '' 현재주는 datepart("ww",now()-1day) '' 우리는 월~일요일까지를 한주로 한다
dim defaultWW : defaultWW=DatePart("ww",dateadd("d",-1,now()))

'' 이번주 
dim thisMon : thisMon = LEFT(dateadd("d",DatePart("w",now())*-1+2,now()),10)
dim thisSun : thisSun = dateadd("d",6,thisMon)

'' 오늘요일
dim thisW : thisW = datepart("w",now())

if (vDategbn="") then vDategbn="O" ''주문일

If vSDate = "" Then
    if (thisW=2 or thisW=3) then  ''월, 화요일은 일주일 전값
        vSDate = dateadd("d",-7,thisMon)
	    vEDate = dateadd("d",-7,thisSun)
    else
	    vSDate = thisMon
	    vEDate = LEFT(date(),10)
    end if
End If


vArrData = getShopifyRows(vArrCols,makerid,itemids)

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script>
function goSearch(){
	
	document.frm1.submit();
}

</script>

<body>
<form name="frm1" method="get" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	브랜드&nbsp; : <% drawSelectBoxDesignerwithName "makerid",makerid %><br>
	상품코드: 
	<textarea id="itemids" rows="5" cols="40" name="itemids"><%=itemids%></textarea>
	최대 500개, 해외가격기준 WSLWEB,  소비가 일단동일
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<% 
dim fld, vArr, rows, cols, col_name, col_wid, col_fmt, col_align, colsplited
%>
<table cellpadding="3" cellspacing="2" border="0" class="a" align="center" width="100%">
<tr bgcolor="#FFFFFF">
    <td width="500">
        
        <% If isArray(vArrCols) Then %>
        <table cellpadding="5" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4" align="center">
        <% For cols = 0 To UBound(vArrCols) %>
            <% 
                colsplited = split(vArrCols(cols),"|") 
                if isArray(colsplited) then
                    col_name = colsplited(0)
                else
                    col_name = colsplited
                end if
            %>
            <td >
                <%=col_name%>
            </td>
        <% Next %>
        </tr>
        <% end if %>
        <% if isArray(vArrData) then %>
        <% For i = 0 To UBound(vArrData,2) %>
        <tr bgcolor="#FFFFFF" align="center">
            <% for cols=0 To UBound(vArrCols) %>
            <%
                colsplited = split(vArrCols(cols),"|") 
                col_fmt = ""
                col_align = ""
                col_wid = ""
                if isArray(colsplited) then
                    if UBOUND(colsplited)>0 then col_fmt = colsplited(1)
                    if UBOUND(colsplited)>1 then col_align = colsplited(2)    
                    if UBOUND(colsplited)>2 then col_wid = colsplited(3) 
                else
                    col_fmt  = "S"
                    col_align = ""
                    col_wid  = ""
                end if
            %>
            <td <%= CHKIIF(col_wid<>"","width='"&col_wid&"'","") %> <%= CHKIIF(col_align="R","align='right'","") %> >
                <% if (col_fmt="N") then %>
                    <%=FormatNumber(vArrData(cols,i),0)%>
                <% else %>
                    <%=vArrData(cols,i)%>
                <% end if %>
            </td>
            <% next %>
        </tr>
        <% next %>
        </table>
        <% else %>
        No data
        <% end if %>
        
    </td>
	
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->