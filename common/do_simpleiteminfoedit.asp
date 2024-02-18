<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품수정
' Hieditor : 서동석 생성
'			 2021.03.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/RackCodeFunction.asp"-->
<%
dim itemid, limityn, limitdispyn, vChangeContents, vSCMChangeSQL
vChangeContents = "- HTTP_REFERER : " & request.ServerVariables("HTTP_REFERER") & vbCrLf
dim dispyn, sellyn, isusing, isextusing
dim pojangok, itemrackcode, subitemrackcode
dim itemoptionarr, optisusingarr
dim optremainnoarr, optrackcodearr, suboptrackcodearr
dim danjongyn
dim blnsellreserve, dsellreservedate,blnSRCondition
dim orgSellyn, orgsellSTDate, settime
dim optdanjongyn, optdanjongynarr
itemid  		= RequestCheckVar(request("itemid"),16)
limityn 		= RequestCheckVar(request("limityn"),1)
limitdispyn 	= RequestCheckVar(request("limitdispyn"),1)
if limitdispyn 	= "" THEN limitdispyn = "Y"
dispyn  		= RequestCheckVar(request("dispyn"),1)
sellyn  		= RequestCheckVar(request("sellyn"),1)
isusing 		= RequestCheckVar(request("isusing"),1)
isextusing   	= RequestCheckVar(request("isextusing"),1)
itemrackcode 	= RequestCheckVar(request("itemrackcode"),8)
subitemrackcode = RequestCheckVar(request("subitemrackcode"),8)
danjongyn    	= RequestCheckVar(request("danjongyn"),1)
settime			= requestCheckvar(Request("settime"),2)
itemoptionarr 	= request("itemoptionarr")
optisusingarr	= request("optisusingarr")
optdanjongynarr	= request("optdanjongynarr")
optremainnoarr  = request("optremainnoarr")
optrackcodearr  = request("optrackcodearr")
suboptrackcodearr  = request("suboptrackcodearr")

itemoptionarr 	= split(itemoptionarr,",")
optisusingarr 	= split(optisusingarr,",")
optdanjongynarr 	= split(optdanjongynarr,",")
optremainnoarr  = split(optremainnoarr,",")
optrackcodearr  = split(optrackcodearr,",")
suboptrackcodearr  = split(suboptrackcodearr,",")

blnsellreserve = requestCheckvar(Request("chkSR"),1)
dsellreservedate = requestCheckvar(Request("dSR"),10)
blnSRCondition= requestCheckvar(Request("chkSRC"),1)
IF blnsellreserve = "" THEN blnsellreserve = "N"
dsellreservedate = dsellreservedate& " " & settime & ":00:00"

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, i

    sqlStr = " select sellyn, sellSTDate FROM db_item.dbo.tbl_item WHERE itemid =" + CStr(itemid)
    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
    	orgSellyn       = rsget("sellyn")
    	orgsellSTDate   = rsget("sellSTDate")
    end if
    rsget.close

	'// 보조 랙코드 입력
	sqlStr = " insert into [db_item].[dbo].[tbl_item_logics_addinfo](itemid,subitemrackcode) "
	sqlStr = sqlStr + " select top 1 i.itemid, i.itemrackcode "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
	sqlStr = sqlStr + " 	left join [db_item].[dbo].[tbl_item_logics_addinfo] l "
	sqlStr = sqlStr + " 	on "
	sqlStr = sqlStr + " 		i.itemid = l.itemid "
	sqlStr = sqlStr + " where i.itemid = " & CStr(itemid) & " and l.itemid is NULL "
	''dbget.Execute sqlStr

if (limityn="Y") then
	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " , isextusing='" + isextusing + "'" + VBCrlf
	sqlStr = sqlStr + " , danjongyn='" + danjongyn + "'" + VBCrlf
	''sqlStr = sqlStr + " , itemrackcode='" + itemrackcode + "'" + VBCrlf
	sqlStr = sqlStr + " , lastupdate=getdate()" + VBCrlf
	sqlStr = sqlStr + " , limitdispyn = '"+limitdispyn+"'"+ VBCrlf
	    if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf
	    end if
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	dbget.Execute sqlStr

    Call RF_SetItemRackCode("10", itemid, itemrackcode)
    Call RF_SetSubItemRackCode("10", itemid, subitemrackcode)

	vChangeContents = vChangeContents & "- 한정판매여부 : limityn = " & limityn & vbCrLf
	vChangeContents = vChangeContents & "- 상품판매여부 : sellyn = " & sellyn & vbCrLf
	vChangeContents = vChangeContents & "- 상품사용여부 : isusing = " & isusing & vbCrLf
	vChangeContents = vChangeContents & "- 제휴사용여부 : isextusing = " & isextusing & vbCrLf
	vChangeContents = vChangeContents & "- 상품단종여부 : danjongyn = " & danjongyn & vbCrLf
	vChangeContents = vChangeContents & "- 한정노출여부 : limitdispyn = " & limitdispyn & vbCrLf

	''옵션한정여부한정
	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbget.Execute sqlStr

	for i=0 to UBound(itemoptionarr)
		if (Len(Trim(itemoptionarr(i)))=4) then
			if (itemoptionarr(i)="0000") then
				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
				sqlStr = sqlStr + " set limitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , limitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

				dbget.Execute sqlStr

				vChangeContents = vChangeContents & "- 0000 옵션한정수량 : limitno = " & optremainnoarr(i) & vbCrLf
			else
				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optdanjongyn='" + optdanjongynarr(i) + "'" + VBCrlf
                sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optlimitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , optlimitsold=" + "0" + "" + VBCrlf
				''sqlStr = sqlStr + " , optrackcode=" + CHKIIF(optrackcodearr(i)="", "NULL", "'" + optrackcodearr(i) + "'") + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"
				dbget.Execute sqlStr

                if (optrackcodearr(i) = "") then
                    Call RF_DelItemRackCodeByOption("10", itemid, itemoptionarr(i))
                else
                    Call RF_SetItemRackCodeByOption("10", itemid, itemoptionarr(i), optrackcodearr(i))
                end if

                if (suboptrackcodearr(i) = "") then
                    Call RF_DelSubItemRackCodeByOption("10", itemid, itemoptionarr(i))
                else
                    Call RF_SetSubItemRackCodeByOption("10", itemid, itemoptionarr(i), suboptrackcodearr(i))
                end if

				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션사용여부 : isusing = " & optisusingarr(i) & vbCrLf
                vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션단종구분 : optdanjongyn = " & optdanjongynarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션판매여부 : optsellyn = " & optisusingarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션한정수량 : optlimitno = " & optremainnoarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션랙코드 : optrackcode = " & optrackcodearr(i) & vbCrLf
                vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션보조랙코드 : suboptrackcode = " & suboptrackcodearr(i) & vbCrLf
			end if
		end if
	next
else

	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " , isextusing='" + isextusing + "'" + VBCrlf
	sqlStr = sqlStr + " , danjongyn='" + danjongyn + "'" + VBCrlf
	''sqlStr = sqlStr + " , itemrackcode='" + itemrackcode + "'" + VBCrlf
	sqlStr = sqlStr + " , lastupdate=getdate()"+ VBCrlf
	sqlStr = sqlStr + " , limitdispyn = '"+limitdispyn+"'"+ VBCrlf
	if orgSellyn <>"Y" and sellyn ="Y" and isNull(orgsellSTDate) then
	sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf
	    end if
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)
	dbget.Execute sqlStr

    Call RF_SetItemRackCode("10", itemid, itemrackcode)
    Call RF_SetSubItemRackCode("10", itemid, subitemrackcode)

	vChangeContents = vChangeContents & "- 한정판매여부 : limityn = " & limityn & vbCrLf
	vChangeContents = vChangeContents & "- 상품판매여부 : sellyn = " & sellyn & vbCrLf
	vChangeContents = vChangeContents & "- 상품사용여부 : isusing = " & isusing & vbCrLf
	vChangeContents = vChangeContents & "- 제휴사용여부 : isextusing = " & isextusing & vbCrLf
	vChangeContents = vChangeContents & "- 상품단종여부 : danjongyn = " & danjongyn & vbCrLf
	vChangeContents = vChangeContents & "- 한정노출여부 : limitdispyn = " & limitdispyn & vbCrLf


	''옵션한정여부한정
	sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbget.Execute sqlStr

	for i=0 to UBound(itemoptionarr)
		if (Len(Trim(itemoptionarr(i)))=4) then
			if (itemoptionarr(i)="0000") then
				sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
				sqlStr = sqlStr + " set limitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , limitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

				dbget.Execute sqlStr

				vChangeContents = vChangeContents & "- 0000 옵션한정수량 : limitno = " & optremainnoarr(i) & vbCrLf
			else
				sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optdanjongyn='" + optdanjongynarr(i) + "'" + VBCrlf
                sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optlimitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , optlimitsold=" + "0" + "" + VBCrlf
				''sqlStr = sqlStr + " , optrackcode=" + CHKIIF(optrackcodearr(i)="", "NULL", "'" + optrackcodearr(i) + "'") + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"
				dbget.Execute sqlStr

                if (optrackcodearr(i) = "") then
                    Call RF_DelItemRackCodeByOption("10", itemid, itemoptionarr(i))
                else
                    Call RF_SetItemRackCodeByOption("10", itemid, itemoptionarr(i), optrackcodearr(i))
                end if

                if (suboptrackcodearr(i) = "") then
                    Call RF_DelSubItemRackCodeByOption("10", itemid, itemoptionarr(i))
                else
                    Call RF_SetSubItemRackCodeByOption("10", itemid, itemoptionarr(i), suboptrackcodearr(i))
                end if

				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션사용여부 : isusing = " & optisusingarr(i) & vbCrLf
                vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션단종구분 : optdanjongyn = " & optdanjongynarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션판매여부 : optsellyn = " & optisusingarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션한정수량 : optlimitno = " & optremainnoarr(i) & vbCrLf
				vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션랙코드 : optrackcode = " & optrackcodearr(i) & vbCrLf
                vChangeContents = vChangeContents & "- " & Trim(itemoptionarr(i)) & " 옵션보조랙코드 : suboptrackcode = " & suboptrackcodearr(i) & vbCrLf
			end if
		end if
	next
end if


''상품옵션수량
sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
sqlStr = sqlStr + " from (" + VBCrlf
sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
sqlStr = sqlStr + " ) T" + VBCrlf
sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf

dbget.Execute sqlStr

''sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
''sqlStr = sqlStr + " set optionname=v.codeview" + VBCrlf
''sqlStr = sqlStr + " from [db_item].[dbo].vw_all_option v" + VBCrlf
''sqlStr = sqlStr + " where  [db_item].[dbo].tbl_item_option.itemid=" + CStr(itemid) + VBCrlf
''sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemoption=v.optioncode" + VBCrlf
''
''dbget.Execute sqlStr


	''상품한정수량
	sqlStr = "update [db_item].[dbo].tbl_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from [db_item].[dbo].tbl_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where [db_item].[dbo].tbl_item.itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " and [db_item].[dbo].tbl_item.optioncnt>0"

	dbget.Execute sqlStr

    sqlStr = " update [db_item].[dbo].tbl_item_option "
    sqlStr = sqlStr + " set optlimityn = T.limityn " ''optsellyn = T.sellyn,
    sqlStr = sqlStr + " from ( "
    sqlStr = sqlStr + "     select top 1 sellyn, limityn from [db_item].[dbo].tbl_item where itemid = " + CStr(itemid) + " "
    sqlStr = sqlStr + " ) T "
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "

    dbget.Execute sqlStr

    '' 한정 판매 0 이면 일시 품절 처리
    sqlStr = " update [db_item].[dbo].tbl_item "
	sqlStr = sqlStr + " set sellyn='S'"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "
	sqlStr = sqlStr + " and sellyn='Y'"
	sqlStr = sqlStr + " and limityn='Y'"
	sqlStr = sqlStr + " and limitno-limitSold<1"

    dbget.Execute sqlStr

'response.write blnsellreserve & "<br>"
'response.write blnSRCondition

'오픈예약 처리
dim objCmd, returnValue
IF blnsellreserve = "Y" and blnSRCondition = "1" THEN '오픈예약 설정여부 + 오픈예약 조건 충족여부(텐배-재고있음, 업체배송)
		 Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
						.ActiveConnection = dbget
						.CommandType = adCmdText
						.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_Insert]("&itemid&",'"&dsellreservedate&"','"&session("ssBctId")&"')}"
						.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
						.Execute, , adExecuteNoRecords
						End With
					    returnValue = objCmd(0).Value
				Set objCmd = nothing
				IF returnValue <>  1  THEN
		        Call Alert_msg("처리중 에러가 발생했습니다. errcode : 오픈예약" )
		        response.end
				END IF
ELSEIF blnsellreserve="N" THEN
			 Set objCmd = Server.CreateObject("ADODB.COMMAND")
					With objCmd
						.ActiveConnection = dbget
						.CommandType = adCmdText
						.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_cancel]("&itemid&",'"&session("ssBctId")&"')}"
						.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
						.Execute, , adExecuteNoRecords
						End With
					    returnValue = objCmd(0).Value
				Set objCmd = nothing
				IF returnValue <>  1  THEN
		        Call Alert_msg("처리중 에러가 발생했습니다. errcode : 오픈예약" )
		        response.end
				END IF
END IF

	vChangeContents = vChangeContents & "- 오픈예약처리 : blnsellreserve = " & blnsellreserve & vbCrLf

	'### 수정 로그 저장(item)
	vSCMChangeSQL = "INSERT INTO [db_log].[dbo].[tbl_scm_change_log](userid, gubun, pk_idx, menupos, contents, refip) "
	vSCMChangeSQL = vSCMChangeSQL & "VALUES('" & session("ssBctId") & "', 'item', '" & itemid & "', '" & Request("menupos") & "', "
	vSCMChangeSQL = vSCMChangeSQL & "'" & vChangeContents & "', '" & Request.ServerVariables("REMOTE_ADDR") & "')"
	dbget.execute(vSCMChangeSQL)
%>
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
