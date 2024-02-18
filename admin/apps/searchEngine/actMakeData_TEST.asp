<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%
''relatedKeyword ¸¸ »ç¿ë

function CheckVaildIP(ref)
    CheckVaildIP = false

    dim VaildIP : VaildIP = Array("172.16.0.106","172.16.0.107","172.16.0.108","172.16.0.109","172.16.0.110","110.93.128.106","110.93.128.107","110.93.128.108","110.93.128.109","110.93.128.110","61.252.133.2", "61.252.133.4", "61.252.133.9","61.252.133.10","61.252.133.80","61.252.133.70","192.168.0.106")
    dim i
    for i=0 to UBound(VaildIP)
        if (VaildIP(i)=ref) then
            CheckVaildIP = true
            exit function
        end if
    next
end function



dim ref : ref = Request.ServerVariables("REMOTE_ADDR")

if (Not CheckVaildIP(ref)) then
	'// TODO : Å×½ºÆ®Áß
     dbget.Close()
     response.end
end if

dim act     : act = requestCheckVar(request("act"),32)
dim param1  : param1 = requestCheckVar(request("param1"),32)
dim sqlStr, i, paramData
dim retCnt : retCnt = 0

dim retval, tmpval2
dim recommandKeyword
dim partKeyword, fullkeyword, partKeywordPrev
dim orgKeyword, relatedKeyword, orgKeywordPrev

select Case act

    Case "recommandKeyword"

		retval = ""

        if (param1<>"") and isNumeric(param1) then
            sqlStr = " select top "&param1&" recommandKeyword "&VbCRLF
        else
		    sqlStr = " select top 15000 recommandKeyword "&VbCRLF
	    end if
		sqlStr = sqlStr + " from db_log.dbo.tbl_keyword_recommand "
		sqlStr = sqlStr + " where recommandKeyword not in ('1','-1','A')" '' eastone Ãß°¡
		sqlStr = sqlStr + " order by searchCount desc "

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
				recommandKeyword = rsget("recommandKeyword")

				''retval = retval & replace(replace(replace(recommandKeyword," ",""),"script",""),"iframe","") & vbCrLf
				retval = retval & TRIM(replace(replace(recommandKeyword,"script",""),"iframe","")) & vbCrLf
				
				rsget.MoveNext
    		loop
        end if
        rsget.close
        
        retval = replace(retval," ","")
		response.write retval
    Case "relatedKeyword"
		'// ¿¬°ü°Ë»ö¾î
        sqlStr = "exec [db_log].[dbo].[sp_Ten_SearchKey_related_HandADD]"
        dbget.Execute sqlStr

		sqlStr = " select top 30000 T.orgKeyword, T.relatedKeyword "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + " 		select orgKeyword, relatedKeyword, searchCount, keywordRank "
		sqlStr = sqlStr + " 		, ROW_NUMBER() OVER (Partition by orgKeyword order by keywordRank, searchCount desc, regdate) as ranky "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 		db_log.dbo.tbl_keyword_related_for_search "
		sqlStr = sqlStr + " 		where keywordRank < 10 "
		sqlStr = sqlStr + " 	) T "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + " 	and T.orgKeyword not in ('1', '-1','2') "
		sqlStr = sqlStr + " 	and T.ranky <= 5 "
		sqlStr = sqlStr + " 	and Len(T.relatedKeyword) > 1"
 	    sqlStr = sqlStr + " 	and Len(T.orgKeyword ) > 1"
 	    sqlStr = sqlStr + " 	and LEFT(T.orgKeyword,1) not in ('?','&','¢¬','¡í','.')"
		sqlStr = sqlStr + " order by "
		sqlStr = sqlStr + " 	T.orgKeyword, T.searchCount desc "

		orgKeyword = ""
		relatedKeyword = ""
		orgKeywordPrev = ""
		retval = ""

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
				''orgKeyword = replace(replace(replace(Replace(rsget("orgKeyword"), ",", "")," ",""),"script",""),"iframe","")
				''relatedKeyword = replace(replace(replace(Replace(rsget("relatedKeyword"), ",", "")," ",""),"script",""),"iframe","")

                orgKeyword = replace(replace(replace(rsget("orgKeyword"), ",", ""),"script",""),"iframe","")
				relatedKeyword = replace(replace(Replace(rsget("relatedKeyword"), ",", ""),"script",""),"iframe","")

				if (orgKeywordPrev <> orgKeyword) then
					if (retval <> "") then
						retval = retval & vbCrLf
					end if

					orgKeywordPrev = orgKeyword
					retval = retval & orgKeyword & "," & relatedKeyword
				else
					retval = retval & "," & relatedKeyword
				end if

				rsget.MoveNext
    		loop
        end if
        rsget.close

		response.write retval
    Case "relatedKeywordORG"
		'// »ç¿ë¾ÈÇÔ
		dbget.Close()
		response.end

		sqlStr = " select T.partKeyword, T.fullkeyword, T.searchCount, T.partSearchCount, T.ranky "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	( "
		sqlStr = sqlStr + " 	select "
		sqlStr = sqlStr + " 		c.recommandKeyword as partKeyword "
		sqlStr = sqlStr + " 		, SF.fullkeyword "
		sqlStr = sqlStr + " 		, SF.searchCount "
		sqlStr = sqlStr + " 		, c.searchCount as partSearchCount "
		sqlStr = sqlStr + " 		, ROW_NUMBER() OVER (Partition by SF.partKeyword order by SF.searchcount desc) as ranky "
		sqlStr = sqlStr + " 	from db_log.dbo.tbl_keyword_recommand c "
		sqlStr = sqlStr + " 	join db_log.dbo.tbl_keyword_recommand_for_search SF "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		c.recommandKeyword = SF.partKeyword "
		sqlStr = sqlStr + " ) T "
		sqlStr = sqlStr + " where T.ranky <= 5 and Len(T.partKeyword) > 1"
		sqlStr = sqlStr + " and T.partKeyword not in ('1','-1')" '' eastone Ãß°¡
		sqlStr = sqlStr + " order by T.partSearchCount desc, T.partKeyword "

		partKeyword = ""
		fullkeyword = ""
		partKeywordPrev = ""
		retval = ""

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
				''partKeyword = replace(replace(replace(Replace(rsget("partKeyword"), ",", "")," ",""),"script",""),"iframe","")
				''fullkeyword = replace(replace(replace(Replace(rsget("fullkeyword"), ",", "")," ",""),"script",""),"iframe","")

                partKeyword = replace(replace(Replace(rsget("partKeyword"), ",", ""),"script",""),"iframe","")
				fullkeyword = replace(replace(Replace(rsget("fullkeyword"), ",", ""),"script",""),"iframe","")
				if (partKeywordPrev <> partKeyword) then
					if (retval <> "") then
						retval = retval & vbCrLf
					end if

					partKeywordPrev = partKeyword
					retval = retval & partKeyword & "," & fullkeyword
				else
					retval = retval & "," & fullkeyword
				end if

				rsget.MoveNext
    		loop
        end if
        rsget.close

		response.write retval
    Case "realTimeKeyword"

        if (param1<>"") and isNumeric(param1) then
			sqlStr = " select top " + CStr(param1) + " currKeyword "
        else
		    sqlStr = " select top 20 currKeyword "
	    end if

		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " db_log.dbo.tbl_keyword_log "
		sqlStr = sqlStr + " where DateDiff(hh, regdate, getdate()) <= 6 "
		sqlStr = sqlStr + " and currKeyword not in ('1','-1')" '' eastone Ãß°¡
		sqlStr = sqlStr + " group by currKeyword "
		sqlStr = sqlStr + " order by count(*) desc "
	'' ---------------------------------------------------------------------------
    '' 2013/04/23 ¼öÁ¤
        if (param1<>"") and isNumeric(param1) then
            sqlStr = " select top " + CStr(param1) + " A.currKeyword "
        else
            sqlStr = " select top 20 A.currKeyword "
        end if
        sqlStr = sqlStr + " from ("
        sqlStr = sqlStr + " select top 2000 currKeyword , count(*) as CNT1 , row_Number() over (order by  count(*) desc) as rank1"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + "  db_log.dbo.tbl_keyword_log "
        sqlStr = sqlStr + "  where DateDiff(hh, regdate, getdate()) <= 6"
        sqlStr = sqlStr + "  and currKeyword not in ('1','-1')"
        sqlStr = sqlStr + "  group by currKeyword "
        sqlStr = sqlStr + "  order by count(*) desc "
        sqlStr = sqlStr + " ) A"
        sqlStr = sqlStr + " left join ("
        sqlStr = sqlStr + "  select top 2000 currKeyword , count(*) as CNT2 , row_Number() over (order by  count(*) desc) as rank2"
        sqlStr = sqlStr + "  from "
        sqlStr = sqlStr + "  db_log.dbo.tbl_keyword_log "
        sqlStr = sqlStr + "  where DateDiff(hh, regdate, getdate()) <= 18 and (DateDiff(hh, regdate, getdate()) >6)"
        sqlStr = sqlStr + "  and currKeyword not in ('1','-1')"
        sqlStr = sqlStr + "  group by currKeyword "
        sqlStr = sqlStr + "  order by count(*) desc "
        sqlStr = sqlStr + " ) B on A.currKeyword=B.currKeyword"
        sqlStr = sqlStr + " where 1=1"
        sqlStr = sqlStr + " and isNULL(B.rank2,2000)-A.rank1>=-1"
        sqlStr = sqlStr + " and isNULL(B.CNT2,0)>1"
        sqlStr = sqlStr + " order by (isNULL(B.rank2,2000)-A.rank1-A.rank1*20+CNT1*2) desc"

	'' ---------------------------------------------------------------------------
    '' 2014/12/12 ¼öÁ¤
        sqlStr = "exec [db_log].[dbo].[sp_Ten_SearchKey_top_HandADD]"
        dbget.Execute sqlStr

        if (param1<>"") and isNumeric(param1) then
            sqlStr = " select top " + CStr(param1) + " topKeyword, (case when rankChange is NULL then ':new' else ':' + convert(varchar, rankChange) end) as rankChange "
        else
            sqlStr = " select top 20 topKeyword, (case when rankChange is NULL then ':new' else ':' + convert(varchar, rankChange) end) as rankChange "
        end if
		sqlStr = sqlStr + " from db_log.dbo.tbl_keyword_top_for_search "
		sqlStr = sqlStr + " where topKeyword not in ('1','-1','','Â¸Â¶Â¸Â®Â¾ÃˆÃ„Ã‰ÃŒÆ® Ã†Ä¿Ã¬Ä¡','A','¹«·á¹è¼Û','Å©¸®½º¹Ú½º','¼ê·œrstxtÄÉÀÌ½º')"
		sqlStr = sqlStr + " order by searchCount desc, topKeyword "

		retval = ""

        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
            do until rsget.eof
				''retval = retval & replace(replace(replace(rsget("topKeyword")," ",""),"script",""),"iframe","") & vbCrLf
				''retval = retval & replace(replace(rsget("topKeyword"),"script",""),"iframe","") & vbCrLf
				retval = retval & replace(replace(rsget("topKeyword"),"script",""),"iframe","") & rsget("rankChange") & vbCrLf
				rsget.MoveNext
    		loop
        end if
        rsget.close

		response.write retval
    Case ELSE
        response.Write "S_ERR|Not Valid - "&act
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
