<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/report/simpleQryCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<%
Dim menuGrpIdx : menuGrpIdx = fnGetMenuGrpIDx()
Dim qryidx : qryidx = requestCheckvar(request("qryidx"),10)
Dim oSimpleQuery, i, j
Dim MaxparamCNT : MaxparamCNT = 10
Dim iparamboxtype, iparamname, iparamBoxhtml, iparamval, iparammaxlen
Dim bparam1, bparam2

' ReDim paramArr(MaxparamCNT)
' for i=0 to MaxparamCNT-1
'     paramArr(i) = requestCheckvar(request("param"&(i+1)),32)
' Next
dim page : page = requestCheckvar(request("page"),10)
if (page="") then page=1

Dim oQryParam , retExeType, vArrCols, vArrData, retReturnValue
Dim retVal
if (qryidx<>"") then
    SET oQryParam = new CSimpleQuery
    oQryParam.FRectQryidx = qryidx
    oQryParam.getQueryParamArr

    for i=0 to oQryParam.FParamCount-1
        iparamname = oQryParam.FQryParamList(i).Fparamname
        oQryParam.FQryParamList(i).FStoredparamVal = oQryParam.FQryParamList(i).getRequestParam ''requestCheckVar(request(iparamname),oQryParam.FQryParamList(i).Fparamlength)

        if (oQryParam.FQryParamList(i).FStoredparamVal="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
            oQryParam.FQryParamList(i).FStoredparamVal = oQryParam.FQryParamList(i).getDefaultVAl
        end if
    next

    retVal = oQryParam.ExecSimpleQuery(retExeType, vArrCols, vArrData, retReturnValue)

end if



''-----------------------------------------
dim vReportType
Dim vSDate, vEDate, vChannel, vOrdType
Dim vDategbn, addparam1, addparam2

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


SET oSimpleQuery = new CSimpleQuery
    oSimpleQuery.FPageSize = 30
    oSimpleQuery.FCurrPage = page

	' oSimpleQuery.FRectParams = vSDate
	' oSimpleQuery.FRectEDate = LEFT(DateAdd("d",1,vEDate),10)
	' oSimpleQuery.FRectDateGbn = vDategbn
	' oSimpleQuery.FRectReportType = vReportType
	' oSimpleQuery.FRectChannel = vChannel
	' oSimpleQuery.FRectAddParam1 = addparam1
	' oSimpleQuery.FRectAddParam2 = addparam2
	' oSimpleQuery.FRectOrderType = vOrdType

	' vArrData = oSimpleQuery.getSimpleReport(vArrCols)

SET oSimpleQuery = nothing

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script>
$(function() {
    <%
    if (qryidx<>"") then
    for i=0 to oQryParam.FParamCount-1
        iparamboxtype = oQryParam.FQryParamList(i).Fparamboxtype
        iparamname = oQryParam.FQryParamList(i).Fparamname
        iparamBoxhtml = ""
        SELECT CASE iparamboxtype
            CASE "yyyymmdd" :
                iparamBoxhtml = "var CAL_"&iparamname&" = new Calendar({"
                iparamBoxhtml = iparamBoxhtml & "inputField : '"&iparamname&"', trigger    : '"&iparamname&"_trigger',"
		        iparamBoxhtml = iparamBoxhtml & " onSelect: function() {"
			    iparamBoxhtml = iparamBoxhtml & "   var date = Calendar.intToDate(this.selection.get());"
			    'iparamBoxhtml = iparamBoxhtml & "   CAL_End.args.min = date;"
			    'iparamBoxhtml = iparamBoxhtml & "   CAL_End.redraw();"
			    iparamBoxhtml = iparamBoxhtml & "   this.hide();"
		        iparamBoxhtml = iparamBoxhtml & "   }, bottomBar: true, dateFormat: '%Y-%m-%d'"
	            iparamBoxhtml = iparamBoxhtml & "});"

            CASE else
                iparamBoxhtml =""
        END SELECT
        response.write iparamBoxhtml&vbCRLF

    Next
    end if
	%>
});

function goSearch(){
	if($("#sdate").val() == ""){
		alert("시작일을 입력하세요");
		return false;
	}
	if($("#edate").val()== ""){
		alert("종료일을 입력하세요");
		return false;
	}
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
    쿼리선택 : <% call drawSimpleQuerySelectBox("qryidx",qryidx,menuGrpIdx) %>
    &nbsp;&nbsp;



    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>
<tr align="center" bgcolor="#F4F4F4">
    <td align="left">
    <% if (qryidx<>"") then %>
    <% for i=0 to oQryParam.FParamCount-1 %>
        <%
        iparamboxtype = oQryParam.FQryParamList(i).Fparamboxtype
        iparamname    = oQryParam.FQryParamList(i).Fparamname
        iparammaxlen  = oQryParam.FQryParamList(i).Fparamlength
        iparamBoxhtml = ""
        iparamval     = ""
        bparam1       = ""
        bparam2       = ""
        SELECT CASE iparamboxtype
            CASE "yyyymmdd" :
                iparamval  = oQryParam.FQryParamList(i).getRequestParam 'requestCheckVar(request(iparamname),iparammaxlen)
                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).getDefaultVAL ''LEFT(dateadd("d",oQryParam.FQryParamList(i).Fdefaultval,now()),10)
                end if
                iparamBoxhtml= oQryParam.FQryParamList(i).FparamTitle & " : <input id='"&iparamname&"' name='"&iparamname&"' value='"&iparamval&"' class='text' size='10' maxlength='10' />"
                iparamBoxhtml = iparamBoxhtml & "<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='"&iparamname&"_trigger' border='0' style='cursor:pointer' align='absmiddle' />"
            CASE "yyyymm" :
                iparamval = oQryParam.FQryParamList(i).getRequestParam

                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).getDefaultVAL ''LEFT(dateadd("m",oQryParam.FQryParamList(i).Fdefaultval,now()),7)
                end if
                bparam1 = LEFT(iparamval,4)
                bparam2 = RIGHT(iparamval,2)

                iparamBoxhtml = ""
                response.write oQryParam.FQryParamList(i).FparamTitle & " : "
                DrawYMBox bparam1,bparam2
                response.write "&nbsp;&nbsp;"

            CASE "box" :
                iparamval  = oQryParam.FQryParamList(i).getRequestParam ''requestCheckVar(request(iparamname),iparammaxlen)

                if (iparamval="") and (oQryParam.FQryParamList(i).Fdefaultval<>"") then
                    iparamval = oQryParam.FQryParamList(i).Fdefaultval
                end if

                iparamBoxhtml = oQryParam.FQryParamList(i).FparamTitle & " : <input type='text' name='"&iparamname&"' value='"&iparamval&"' size='"&iparammaxlen&"' maxlength='"&iparammaxlen&"'>"

            CASE else
                iparamBoxhtml = "?"&iparamboxtype&"?"

        END SELECT

        if (iparamBoxhtml<>"") then iparamBoxhtml=iparamBoxhtml&"&nbsp;&nbsp;"
        response.write iparamBoxhtml
        %>

    <% next %>
    <% end if %>
    </td>
</tr>

</table>
</form>
<p>
    <% 'Call drawReportDescription(vReportType) %>
</p>
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
                    <% if vArrData(cols,i)="" or isnull(vArrData(cols,i)) then %>
                        0
                    <% else %>
                        <%=FormatNumber(vArrData(cols,i),0)%>
                    <% end if %>
                <% else %>
                    <%=vArrData(cols,i)%>
                <% end if %>
            </td>
            <% next %>
        </tr>
        <%
                if (i mod 500)=499 then
                    response.flush
                end if
            next
        %>
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
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
