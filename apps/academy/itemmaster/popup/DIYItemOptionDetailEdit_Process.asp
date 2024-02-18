<%@ codepage="65001" language=vbscript %>
<% option explicit %>
<%
response.Charset="UTF-8"
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/apps/academy/itemmaster/itemOptionLib.asp"-->
<!-- #include virtual="/apps/academy/lib/chkItem.asp"-->
<%
dim itemid, limityn, makerid
dim dispyn, sellyn, isusing
dim itemoptionarr, optisusingarr
dim optremainnoarr, TotalOptLimitNo

itemid  = requestCheckVar(request("itemid"),10)
makerid = request.cookies("partner")("userid")

If (ItemCheckMyItemYN(makerid,itemid)<>true) Then
	Response.Write "<script>alert('상품이 없거나 잘못된 상품입니다.');</script>"
	Response.end
End If
limityn = requestCheckVar(request("limityn"),10)
dispyn  = "Y"
sellyn  = "Y"
isusing = "Y"

itemoptionarr 	= request("itemoptionarr")
optisusingarr	= request("optisusingarr")
optremainnoarr  = request("optremainnoarr")

itemoptionarr 	= split(itemoptionarr,",")
optisusingarr 	= split(optisusingarr,",")
optremainnoarr     = split(optremainnoarr,",")

dim refer
refer = request.ServerVariables("HTTP_REFERER")
TotalOptLimitNo=0
dim sqlStr, i

if (limityn="Y") then
	sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " , lastupdate=getdate()"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbACADEMYget.Execute sqlStr

	''옵션한정여부한정
	sqlStr = "update db_academy.dbo.tbl_diy_item_option" + VBCrlf
	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbACADEMYget.Execute sqlStr
	Dim intch
	for i=0 to UBound(itemoptionarr)
		if (Len(Trim(itemoptionarr(i)))=4) then
			if (itemoptionarr(i)="0000") then
				sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
				sqlStr = sqlStr + " set limitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , limitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

				dbACADEMYget.Execute sqlStr
			else
				sqlStr = "update db_academy.dbo.tbl_diy_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optlimitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , optlimitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"

				dbACADEMYget.Execute sqlStr
			end If
			intch = optremainnoarr(i)
		end If
		'Response.write Trim(optremainnoarr(i)) & "<br>"
		TotalOptLimitNo = TotalOptLimitNo + Cint(intch)
	Next
	'Response.end
else
	sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
	sqlStr = sqlStr + " set limityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " , sellyn='" + sellyn + "'" + VBCrlf
	sqlStr = sqlStr + " , isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " , lastupdate=getdate()"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbACADEMYget.Execute sqlStr


	''옵션한정여부한정
	sqlStr = "update db_academy.dbo.tbl_diy_item_option" + VBCrlf
	sqlStr = sqlStr + " set optlimityn='" + limityn + "'" + VBCrlf
	sqlStr = sqlStr + " where itemid=" + CStr(itemid)

	dbACADEMYget.Execute sqlStr

	for i=0 to UBound(itemoptionarr)
		if (Len(Trim(itemoptionarr(i)))=4) then
			if (itemoptionarr(i)="0000") then
				sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
				sqlStr = sqlStr + " set limitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , limitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)

				dbACADEMYget.Execute sqlStr
			else
				sqlStr = "update db_academy.dbo.tbl_diy_item_option" + VBCrlf
				sqlStr = sqlStr + " set isusing='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optsellyn='" + optisusingarr(i) + "'" + VBCrlf
				sqlStr = sqlStr + " , optlimitno=" + optremainnoarr(i) + "" + VBCrlf
				sqlStr = sqlStr + " , optlimitsold=" + "0" + "" + VBCrlf
				sqlStr = sqlStr + " where itemid=" + CStr(itemid)
				sqlStr = sqlStr + " and itemoption='" + Trim(itemoptionarr(i)) + "'"

				dbACADEMYget.Execute sqlStr
			end if
		end if
	Next
	TotalOptLimitNo=0
end if


''상품옵션수량
sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
sqlStr = sqlStr + " set optioncnt=IsNULL(T.optioncnt,0)" + VBCrlf
sqlStr = sqlStr + " from (" + VBCrlf
sqlStr = sqlStr + " 	select count(itemoption) as optioncnt" + VBCrlf
sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_item_option" + VBCrlf
sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
sqlStr = sqlStr + " ) T" + VBCrlf
sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item.itemid=" + CStr(itemid) + VBCrlf

dbACADEMYget.Execute sqlStr

	''상품한정수량
	sqlStr = "update db_academy.dbo.tbl_diy_item" + VBCrlf
	sqlStr = sqlStr + " set limitno=IsNULL(T.optlimitno,0), limitsold=IsNULL(T.optlimitsold,0)" + VBCrlf
	sqlStr = sqlStr + " from (" + VBCrlf
	sqlStr = sqlStr + " 	select sum(optlimitno) as optlimitno, sum(optlimitsold) as optlimitsold" + VBCrlf
	sqlStr = sqlStr + " 	from db_academy.dbo.tbl_diy_item_option" + VBCrlf
	sqlStr = sqlStr + " 	where itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " 	and isusing='Y'" + VBCrlf
	sqlStr = sqlStr + " ) T" + VBCrlf
	sqlStr = sqlStr + " where db_academy.dbo.tbl_diy_item.itemid=" + CStr(itemid) + VBCrlf
	sqlStr = sqlStr + " and db_academy.dbo.tbl_diy_item.optioncnt>0"

	dbACADEMYget.Execute sqlStr

    sqlStr = " update db_academy.dbo.tbl_diy_item_option "
    sqlStr = sqlStr + " set optlimityn = T.limityn " ''optsellyn = T.sellyn,
    sqlStr = sqlStr + " from ( "
    sqlStr = sqlStr + "     select top 1 sellyn, limityn from db_academy.dbo.tbl_diy_item where itemid = " + CStr(itemid) + " "
    sqlStr = sqlStr + " ) T "
    sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " "
    
    dbACADEMYget.Execute sqlStr
    
    '' 한정 판매 0 이면 일시 품절 처리
    sqlStr = " update db_academy.dbo.tbl_diy_item "
	sqlStr = sqlStr + " set sellyn='S'"
	sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "
	sqlStr = sqlStr + " and sellyn='Y'"
	sqlStr = sqlStr + " and limityn='Y'"
	sqlStr = sqlStr + " and limitno-limitSold<1"
	
    dbACADEMYget.Execute sqlStr
%>
<script>
<!--
	parent.fnOptionStateEditEnd("<%=TotalOptLimitNo%>");
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->