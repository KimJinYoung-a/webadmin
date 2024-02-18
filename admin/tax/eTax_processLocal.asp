<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/EseroTaxCls.asp"-->
<%
Dim mode        : mode =  requestCheckvar(request("mode"),32)
Dim taxKey      : taxKey =  requestCheckvar(request("taxKey"),24)

Dim sqlStr, pCNT, AssignedRow
Dim paramInfo, retParamInfo, RetErr, retErrStr,retErpLinkType

IF (mode="MapByTaxKey") then

    rw taxKey

    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
        ,Array("@taxKey"	,adVarchar, adParamInput,24, taxKey) _
        ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
	)

    sqlStr = "db_partner.dbo.sp_Ten_Esero_Tax_MatchOne_ByTaxKey"

    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

    RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ� or IDX
    retErrStr    = GetValue(retParamInfo, "@retErrStr")   ' ���� �޼���

    IF (RetErr<1) then
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    ELSE
        rw "����:"&RetErr
    END IF
ELSE
    response.write "mode=["&mode&"] ������"
ENd IF

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
