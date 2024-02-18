<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx,mode,itemid, itemoption
dim sellyn,limityn
dim rejectstr
dim optioncnt
dim limitSetno '', limitno, limitsold

idx = request("idx")
mode = request("mode")
itemid = request("itemid")
itemoption = request("itemoption")
sellyn = request("sellyn")
limityn = request("limityn")
limitSetno = request("limitSetno")
rejectstr= html2db(request("rejectstr"))

'limitno = request("limitno")
'limitsold = request("limitsold")

if ((itemoption = "") or (isnull(itemoption) = true)) then
    itemoption = "0000"
end if


dim sqlStr

if (mode="del") then
	sqlStr = "update [db_temp].[dbo].tbl_upche_itemedit" + vbCrlf
	sqlStr = sqlStr + " set isfinish='D'," + vbCrlf
	sqlStr = sqlStr + " rejectstr='" + html2db(rejectstr) + "'" + vbCrlf
	sqlStr = sqlStr + " where idx=" + idx
	
	dbget.Execute sqlStr
else
	dim orgSellyn,orgsellSTDate
 	sqlStr = " select sellyn, sellSTDate FROM db_item.dbo.tbl_item WHERE itemid =" & CStr(itemid)  
  rsget.Open sqlStr,dbget,1
  if Not rsget.Eof then
  	orgSellyn       = rsget("sellyn") 
  	orgsellSTDate   = rsget("sellSTDate") 
  end if
  rsget.close
  
	if (itemoption = "0000") or (itemoption = "XXXX") then
	''옵션이 없는경우
    	sqlStr = "update [db_item].[dbo].tbl_item" + vbCrlf
    	sqlStr = sqlStr + " set sellyn='" + sellyn + "'" + vbCrlf
    	sqlStr = sqlStr + " ,limityn='" + limityn + "'" + vbCrlf
    	if (limitSetno<>"") and (limityn="Y") then
    	    sqlStr = sqlStr + " ,limitno=" + limitSetno + "" + vbCrlf
    	    sqlStr = sqlStr + " ,limitsold=0" + vbCrlf
    	end if
    	sqlStr = sqlStr + " ,lastupdate=getdate()"
    	if orgSellyn <>"Y" and sellyn  ="Y" and isNull(orgsellSTDate) then
	    sqlStr = sqlStr + " , sellSTDate = getdate() "+ VBCrlf        
	    end if
    	sqlStr = sqlStr + " where itemid=" + itemid
    	
    	dbget.Execute sqlStr
    	
    	sqlStr = "update [db_temp].[dbo].tbl_upche_itemedit" + vbCrlf
    	sqlStr = sqlStr + " set isfinish='Y'" + vbCrlf
    	sqlStr = sqlStr + " where idx=" + idx
    	
    	dbget.Execute sqlStr
    	
    else
        'TODO : dispyn 을 isusing 으로 사용한다
        '옵션별로 한정 설정 할 수 없음.: 옵션 한정-> 전체 한정 설정.
        
       ' if sellyn<>"Y" then sellyn="N"
        
        sqlStr = " update [db_item].[dbo].tbl_item_option "
        sqlStr = sqlStr + " set isusing = '" + sellyn + "' "
        sqlStr = sqlStr + " ,optsellyn='" + sellyn + "' "
        if (limitSetno<>"") and (limityn="Y") then
            sqlStr = sqlStr + " ,optlimitno = " + CStr(limitSetno) + " "
            sqlStr = sqlStr + " ,optlimitsold = 0 "
        end if
        sqlStr = sqlStr + " where itemid = " + CStr(itemid) + " and itemoption = '" + Trim(itemoption) + "' "

        dbget.Execute sqlStr

    	sqlStr = "update [db_temp].[dbo].tbl_upche_itemedit" + vbCrlf
    	sqlStr = sqlStr + " set isfinish='Y'" + vbCrlf
    	sqlStr = sqlStr + " where idx=" + idx
    	
    	dbget.Execute sqlStr
    end if

    optioncnt = 0
    sqlStr = " select top 1 optioncnt from [db_item].[dbo].tbl_item "
    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "
	rsget.Open sqlStr, dbget, 1
	if Not rsget.Eof then
		optioncnt = rsget("optioncnt")
	end if
	rsget.close

    if (optioncnt > 0) then
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
		
		dbget.Execute sqlStr
		
		''옵션 갯수 일치
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
        
        ''
        sqlStr = "update [db_item].[dbo].tbl_item_option" + VBCrlf
        sqlStr = sqlStr + " set optlimitYn=i.limityn"  + VBCrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item i"  + VBCrlf
        sqlStr = sqlStr + " where [db_item].[dbo].tbl_item_option.itemid=" + CStr(itemid) + VBCrlf
        sqlStr = sqlStr + " and [db_item].[dbo].tbl_item_option.itemid=i.itemid"  + VBCrlf
        sqlStr = sqlStr + " and i.limityn<>[db_item].[dbo].tbl_item_option.optlimitYn"  + VBCrlf
        
        dbget.Execute sqlStr
    end if
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
<% if (mode="del") then %>
alert('요청 거부 되었습니다.');
location.replace('<%= refer %>');
<% else %>
alert('승인 수정 되었습니다.');
location.replace('<%= refer %>');
<% end if %>
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->