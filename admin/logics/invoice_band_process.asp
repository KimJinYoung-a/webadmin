<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 송장대역관리
' Hieditor : 2021.04.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/invoice_band_cls.asp"-->
<%
dim menupos, i, mode, sqlStr,adminid, refer, basicsongjangcount
dim iidx,siteseq,gubuncd,startsongjangno,endsongjangno,startrealsongjangno,endrealsongjangno
dim remainsongjangcount,basicsongjangyn,isusing, songjangcount, resultcount, songjangdiv
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	mode = requestcheckvar(trim(request("mode")),32)
    iidx = requestcheckvar(getNumeric(trim(request("iidx"))),10)
    siteseq = requestcheckvar(getNumeric(trim(request("siteseq"))),10)
    gubuncd = requestcheckvar(trim(request("gubuncd")),3)
    startsongjangno = requestcheckvar(trim(request("startsongjangno")),12)
    endsongjangno = requestcheckvar(trim(request("endsongjangno")),12)
    startrealsongjangno = requestcheckvar(trim(request("startrealsongjangno")),12)
    endrealsongjangno = requestcheckvar(trim(request("endrealsongjangno")),12)
    basicsongjangyn = requestcheckvar(trim(request("basicsongjangyn")),1)
    isusing = requestcheckvar(trim(request("isusing")),1)
    songjangdiv = requestcheckvar(trim(request("songjangdiv")),32)

resultcount = 0
songjangcount=0
basicsongjangcount=0
if remainsongjangcount="" or isnull(remainsongjangcount) then remainsongjangcount=0

adminid=session("ssBctId")
refer = request.ServerVariables("HTTP_REFERER")

if siteseq="" or isnull(siteseq) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('업체를 선택하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if
if gubuncd="" or isnull(gubuncd) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('출고구분을 선택하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if
if startsongjangno="" or isnull(startsongjangno) or endsongjangno="" or isnull(endsongjangno) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('송장번호(검증키포함)를 입력하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if
if startrealsongjangno="" or isnull(startrealsongjangno) or endrealsongjangno="" or isnull(endrealsongjangno) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('실제송장번호를 입력하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if
if basicsongjangyn="" or isnull(basicsongjangyn) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('기본송장여부를 선택하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if
if isusing="" or isnull(isusing) then
    response.write "<script type='text/javascript'>"
    response.write "    alert('사용여부를 선택하세요.');"
    response.write "    location.replace('"& refer &"');"
    response.write "</script>"
    dbget.close() : dbget_Logistics.close() : response.end
end if

'//신규저장
if mode = "add" then
    sqlStr = "select count(iidx) as cnt"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)"
    sqlStr = sqlStr & " where isusing='Y' and siteseq="& siteseq &""
    sqlStr = sqlStr & " and gubuncd='"& gubuncd &"'"
    sqlStr = sqlStr & " and startrealsongjangno='"& startrealsongjangno &"'"
    sqlStr = sqlStr & " and songjangdiv='"& songjangdiv &"'"

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

    if not rsget_Logistics.EOF then
        songjangcount=rsget_Logistics("cnt")
    end if
    rsget_Logistics.close

    if songjangcount>0 then
        response.write "<script type='text/javascript'>"
        response.write "    alert('이미 등록된 송장번호 입니다.');"
        response.write "    location.replace('"& refer &"');"
        response.write "</script>"
        dbget.close() : dbget_Logistics.close() : response.end
    end if

    if basicsongjangyn="Y" then
        sqlStr = "update [db_aLogistics].[dbo].[tbl_invoice_band] set basicsongjangyn=N'N' where isusing='Y' and siteseq="& siteseq &" and gubuncd='"& gubuncd &"' and songjangdiv='"& songjangdiv &"'"

        'response.write sqlStr & "<br>"
        dbget_Logistics.execute sqlStr
    end if

	sqlStr = "insert into [db_aLogistics].[dbo].[tbl_invoice_band] (" & vbcrlf
	sqlStr = sqlStr & " siteseq,gubuncd,startsongjangno,endsongjangno,startrealsongjangno,endrealsongjangno" & vbcrlf
    sqlStr = sqlStr & " ,basicsongjangyn,isusing,regdate,lastupdate,reguserid,lastuserid, songjangdiv, remainsongjangcount) values (" & vbcrlf
	sqlStr = sqlStr & " "& siteseq &", N'"& gubuncd &"', N'"& startsongjangno &"', N'"& endsongjangno &"', N'"& startrealsongjangno &"'" & vbcrlf
    sqlStr = sqlStr & " , N'"& endrealsongjangno &"', N'"& basicsongjangyn &"', N'"& isusing &"', getdate(), getdate()" & vbcrlf
    sqlStr = sqlStr & " , N'"& adminid &"', N'"& adminid &"', N'"& songjangdiv &"', N'"& endrealsongjangno-startrealsongjangno &"')" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr, resultcount

    iidx=""
    sqlStr ="select SCOPE_IDENTITY() "

    rsget_Logistics.CursorLocation = adUseClient
    rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly
	IF not rsget_Logistics.Eof Then
		iidx = rsget_Logistics(0)
	End IF
	rsget_Logistics.close

    ' 남은송장수 초기값 셋팅
    sqlStr = sqlStr & " update b set remainsongjangcount=convert(numeric(18, 0),endrealsongjangno)-convert(numeric(18, 0),startrealsongjangno)"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)"
    sqlStr = sqlStr & " where remainsongjangcount=0 and convert(numeric(18, 0),startrealsongjangno)<>0 and convert(numeric(18, 0),endrealsongjangno)<>0 and iidx = "&iidx&"" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr

    sqlStr = "declare @startrealsongjangno nvarchar(12)"
    sqlStr = sqlStr & "	    set @startrealsongjangno='';"
    sqlStr = sqlStr & " declare @endrealsongjangno nvarchar(12)"
    sqlStr = sqlStr & "	    set @endrealsongjangno='';"
    sqlStr = sqlStr & " declare @currentsongjangno nvarchar(12)"
    sqlStr = sqlStr & "	    set @currentsongjangno='';"
    sqlStr = sqlStr & " declare @remainsongjangcount bigint"
    sqlStr = sqlStr & "	    set @remainsongjangcount=0;"

    sqlStr = sqlStr & " set @startrealsongjangno='';"
    sqlStr = sqlStr & " set @endrealsongjangno='';"
    sqlStr = sqlStr & " set @currentsongjangno='';"
    sqlStr = sqlStr & " set @remainsongjangcount=0;"

    sqlStr = sqlStr & " select"
    sqlStr = sqlStr & " @currentsongjangno = endrealsongjangno, @remainsongjangcount=convert(bigint,endrealsongjangno)-convert(bigint,startrealsongjangno)"
    sqlStr = sqlStr & " , @startrealsongjangno = startrealsongjangno, @endrealsongjangno = endrealsongjangno"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] with (nolock)"
    sqlStr = sqlStr & " where isusing='Y' and siteseq="& siteseq &" and gubuncd='"& gubuncd &"' and basicsongjangyn='Y' and songjangdiv = '"& songjangdiv &"'"

    sqlStr = sqlStr & " select top 1"
    sqlStr = sqlStr & " @remainsongjangcount=convert(bigint,@currentsongjangno)-REALSONGJANGNO"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_Logistics_songjang_log] with (nolock)"
    sqlStr = sqlStr & " where SiteSEQ = "& siteseq &" and gubuncd='"& gubuncd &"' and DIV_CD = '"& songjangdiv &"'"
    sqlStr = sqlStr & " and REALSONGJANGNO>=@startrealsongjangno and REALSONGJANGNO<=@endrealsongjangno"
    sqlStr = sqlStr & " order by idx desc"

    sqlStr = sqlStr & " update b set remainsongjangcount=@remainsongjangcount"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)"
    sqlStr = sqlStr & " where isusing='Y' and siteseq="& siteseq &" and gubuncd='"& gubuncd &"' and basicsongjangyn='Y' and songjangdiv = '"& songjangdiv &"'"

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr

'//수정
elseif mode = "edit" then
    if iidx="" or isnull(iidx) then
        response.write "<script type='text/javascript'>"
        response.write "    alert('번호를 선택하세요.');"
        response.write "    location.replace('"& refer &"');"
        response.write "</script>"
        dbget.close() : dbget_Logistics.close() : response.end
    end if

    if isusing="N" then
        if basicsongjangyn="Y" then
            response.write "<script type='text/javascript'>"
            response.write "    alert('기본송장으로 사용중인 송장대역은 사용안함 처리 하실수 없습니다.');"
            response.write "    location.replace('"& refer &"');"
            response.write "</script>"
            dbget.close() : dbget_Logistics.close() : response.end
        end if
    end if

    if basicsongjangyn="Y" then
        sqlStr = "update [db_aLogistics].[dbo].[tbl_invoice_band] set basicsongjangyn=N'N' where isusing='Y' and siteseq="& siteseq &" and gubuncd='"& gubuncd &"' and iidx <> "&iidx&" and songjangdiv='"& songjangdiv &"'"

        'response.write sqlStr & "<br>"
        dbget_Logistics.execute sqlStr
    else
        sqlStr = "select count(iidx) as cnt"
        sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)"
        sqlStr = sqlStr & " where isusing='Y' and basicsongjangyn='Y' and siteseq="& siteseq &""
        sqlStr = sqlStr & " and gubuncd='"& gubuncd &"'"
        sqlStr = sqlStr & " and iidx <> "&iidx&""
        sqlStr = sqlStr & " and songjangdiv='"& songjangdiv &"'"

        rsget_Logistics.CursorLocation = adUseClient
        rsget_Logistics.Open sqlStr, dbget_Logistics, adOpenForwardOnly, adLockReadOnly

        if not rsget_Logistics.EOF then
            basicsongjangcount=rsget_Logistics("cnt")
        end if
        rsget_Logistics.close

        if basicsongjangcount<1 then
            response.write "<script type='text/javascript'>"
            response.write "    alert('기본송장여부를 N으로 수정은 불가합니다.\n\n기본송장이 1개 이상 존재 해야 합니다.\n\n사용하실 송장대역에서 기본송장여부를 Y 로 해주세요.');"
            response.write "    location.replace('"& refer &"');"
            response.write "</script>"
            dbget.close() : dbget_Logistics.close() : response.end
        end if
    end if

	sqlStr = "update [db_aLogistics].[dbo].[tbl_invoice_band] set" & vbcrlf
	sqlStr = sqlStr & " siteseq="& siteseq &"" & vbcrlf
    sqlStr = sqlStr & " , gubuncd=N'"& gubuncd &"'" & vbcrlf
    sqlStr = sqlStr & " , startsongjangno=N'"& startsongjangno &"'" & vbcrlf
    sqlStr = sqlStr & " , endsongjangno=N'"& endsongjangno &"'" & vbcrlf
    sqlStr = sqlStr & " , startrealsongjangno=N'"& startrealsongjangno &"'" & vbcrlf
    sqlStr = sqlStr & " , endrealsongjangno=N'"& endrealsongjangno &"'" & vbcrlf
    sqlStr = sqlStr & " , basicsongjangyn=N'"& basicsongjangyn &"'" & vbcrlf
    sqlStr = sqlStr & " , isusing=N'"& isusing &"'" & vbcrlf
    sqlStr = sqlStr & " , lastupdate=getdate()" & vbcrlf
    sqlStr = sqlStr & " , lastuserid=N'"& adminid &"'" & vbcrlf
    sqlStr = sqlStr & " , songjangdiv=N'"& songjangdiv &"'" & vbcrlf
    sqlStr = sqlStr & " where iidx = "&iidx&"" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr, resultcount

    ' 남은송장수 초기값 셋팅
    sqlStr = sqlStr & " update b set remainsongjangcount=convert(numeric(18, 0),endrealsongjangno)-convert(numeric(18, 0),startrealsongjangno)"
    sqlStr = sqlStr & " from [db_aLogistics].[dbo].[tbl_invoice_band] b with (nolock)"
    sqlStr = sqlStr & " where remainsongjangcount=0 and convert(numeric(18, 0),startrealsongjangno)<>0 and convert(numeric(18, 0),endrealsongjangno)<>0 and iidx = "&iidx&"" & vbcrlf

    'response.write sqlStr & "<br>"
    dbget_Logistics.execute sqlStr
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->

<script type='text/javascript'>
    alert('<%= resultcount %>건 저장 되었습니다.');
	opener.location.reload();
	self.close();
</script>
