<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemid, limityn
dim dispyn, sellyn, isusing
dim itemoptionarr, optisusingarr
dim optremainnoarr

itemid  = requestCheckVar(request("itemid"),10)
limityn = requestCheckVar(request("limityn"),10)
dispyn  = requestCheckVar(request("dispyn"),10)
sellyn  = requestCheckVar(request("sellyn"),10)
isusing = requestCheckVar(request("isusing"),10)

itemoptionarr 	= request("itemoptionarr")
optisusingarr	= request("optisusingarr")
optremainnoarr  = request("optremainnoarr")

itemoptionarr 	= split(itemoptionarr,",")
optisusingarr 	= split(optisusingarr,",")
optremainnoarr     = split(optremainnoarr,",")

dim refer
refer = request.ServerVariables("HTTP_REFERER")

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
	next
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
	next
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
<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->