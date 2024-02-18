<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% 
    Response.Charset="UTF-8" 
    Response.ContentType = "application/json"
    Session.CodePage="65001"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/util/base64New.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp"-->
<!-- #include virtual="/lib/classes/admin/PartInfoCls.asp" -->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'###############################################
' PageName : /datagridsample/statistic/emailcustomerlist/datajson.asp
' Discription : datagridsample - emailcustomerlist
' Response : response > 결과
' History : 2019.07.29 
'###############################################

'//헤더 출력


' response.write Request.ServerVariables("request_method") &"<Br/>"
' response.write Request.ServerVariables("query_string")

dim oJson
dim omd , i
dim page
dim skip : skip = request("skip")
dim pageSize : pageSize = request("take")
dim orderby : orderby = request("orderby")
dim searchKey , searchString
dim mode : mode = Request.ServerVariables("request_method")

'// 전송결과 파징
'on Error Resume Next

if skip <> "" then 
    page = cint(skip / pageSize) + 1
else 
    page = 1
end if 

if pageSize = "" then pageSize = 0

'// json객체 선언
SET oJson = jsObject()

IF (Err) then
	oJson("response") = getErrMsg("9999",sFDesc)
	oJson("faildesc") = "처리중 오류가 발생했습니다.1"
else
    dim part_sn , part_name , part_sort , part_isDel , totalCount

    if mode = "GET" then '// 목록
        set omd = New CPart
            omd.FCurrPage = page
            omd.FPageSize = pageSize
            omd.FRectsearchKey = searchKey
            omd.FRectsearchString = searchString
            omd.FRectOrderBy = orderby
            omd.GetPartInfoList

            totalcount = omd.FTotalCount

            oJson("totalCount") = totalcount '// totalcount
            Set oJson("items") = jsArray()

            if omd.FResultCount > 0 then
                ReDim contents_object(omd.FResultCount-1)
                FOR i=0 to omd.FResultCount-1
                    part_sn                 = omd.FItemList(i).Fpart_sn
                    part_name               = omd.FItemList(i).Fpart_name
                    part_sort               = omd.FItemList(i).Fpart_sort
                    part_isDel              = omd.FItemList(i).Fpart_isDel

                    Set oJson("items")(null) = jsObject()
                        oJson("items")(null)("part_sn")      = part_sn
                        oJson("items")(null)("part_name")	 = part_name
                        oJson("items")(null)("part_sort")	 = part_sort
                        oJson("items")(null)("part_isDel")	 = part_isDel
                next
            end if 

        set omd = Nothing

    ELSEIF mode = "POST" THEN '// 입력 / 수정 / 삭제
        dim requestData : requestData = request.form() '// 넘어온 form data.
        
        dim oResult , subMode '// json 파싱 후 처리
        set oResult = JSON.parse(requestData)
            subMode = oResult.mode

        if subMode = "POST" THEN 

            set omd = New CPart
                omd.FRectPartName = oResult.part_name
                omd.FRectPartSortingNumber = oResult.part_sort
                omd.PostPartInfo()
            set omd = Nothing

        ELSEIF subMode = "PUT" THEN 

            set omd = New CPart
                omd.FRectPartNumber = oResult.part_sn
                omd.FRectPartName = oResult.part_name
                omd.FRectPartSortingNumber = oResult.part_sort
                omd.PutPartInfo()
            set omd = Nothing

        ELSEIF subMode = "DELETE" THEN

            set omd = New CPart
                omd.FRectPartNumber = oResult.part_sn
                omd.DeletePartInfo()
            set omd = Nothing

        END IF 

        set oResult = Nothing
    END IF

	'// 결과 출력
	IF (Err) then
		oJson("response") = getErrMsg("9999",sFDesc)
		oJson("faildesc") = "처리중 오류가 발생했습니다.2"
	end if
end if

'Json 출력(JSON)
oJson.flush
Set oJson = Nothing

if ERR then Call OnErrNoti()
'On Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->