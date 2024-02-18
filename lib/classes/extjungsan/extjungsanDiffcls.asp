<%

Class CextJungsanMappingItem

    public Fyyyymm
    public Fsitename
    public FpayPrice
    public Fitemtype
    public Ften_meachul_yyyymm0
    public Ften_deliver_meachul_yyyymm0
    public Ften_jungsan_yyyymm0
    public Fext_jungsan_yyyymm0

    public Ften_meachul_yyyymm1
    public Ften_deliver_meachul_yyyymm1
    public Ften_jungsan_yyyymm1
    public Fext_jungsan_yyyymm1

    public Ften_meachul_yyyymm2
    public Ften_deliver_meachul_yyyymm2
    public Ften_jungsan_yyyymm2
    public Fext_jungsan_yyyymm2

    public Ften_meachul_yyyymm3
    public Ften_deliver_meachul_yyyymm3
    public Ften_jungsan_yyyymm3
    public Fext_jungsan_yyyymm3

    public Ften_meachul_null
    public Ften_deliver_meachul_null
    public Ften_jungsan_null
    public Fext_jungsan_null

    public Ften_meachul_sum
    public Ften_deliver_meachul_sum
    public Ften_jungsan_sum
    public Fext_jungsan_sum

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub
end Class

Class CextJungsanMapping
	public FItemList()

    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectOutMall

    public FRectyyyymm
    public FRectPyyyymm
    public FRectNyyyymm

    public FdiffDate
    public Fmindate
    public Fmaxdate

    public FRectStyyyymm
    public FRectEdyyyymm

    public Function fnGetextMatchingMaster_V2
        dim strSql
        dim minscmdate, maxscmdate, minomdate, maxomdate, mindate, maxdate
        dim minyyyymm, maxyyyymm
        dim i

	    strSql = " SELECT min(case when scmjsdate < omjsdate then scmjsdate else omjsdate end) as mindate, max(case when scmjsdate > omjsdate then scmjsdate else omjsdate end) as maxdate "
	    strSql = strSql & " FROM db_statistics.dbo.tbl_extsite_orderMatching_master "
	    strSql= strSql & " where scmactdate >='"&FRectStyyyymm&"' and scmactdate <='"&FRectEdyyyymm&"'"
	    if FRectOutMall <> "" then
	        strSql= strSql & " and sitename ='"&FRectOutMall&"'"
	    end if
	    ''response.write strSql&"<br>"
	    rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			mindate = rsSTSget("mindate")
			maxdate = rsSTSget("maxdate")

			if isNull(mindate) or isNull(maxdate) then
				FdiffDate = -1
			else
				FdiffDate = datediff("m",mindate,maxdate)
			end if
		else
			FdiffDate = -1
		END IF

		Fmindate = mindate
		Fmaxdate = maxdate
	    rsSTSget.close

        minyyyymm = Left(Fmindate, 7)
        maxyyyymm = Left(Fmaxdate, 7)

        strSql = " select * " & vbCrLf
        strSql = strSql & " from " & vbCrLf
        strSql = strSql & " 	( " & vbCrLf
        strSql = strSql & " 		select " & vbCrLf
        strSql = strSql & " 			m.scmactdate as yyyymm, m.sitename " & vbCrLf
        strSql = strSql & " 			, isNull((select  sum(scmmeachul) from tbl_extsite_orderMatching_master tt where m.scmactdate = tt.scmactdate and m.sitename = tt.sitename),0) as payPrice " & vbCrLf
        strSql = strSql & " 			, '상품' as itemtype " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul else 0 end) as ext_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul else 0 end) as ext_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul else 0 end) as ext_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul else 0 end) as ext_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.ommeachul else 0 end) as ext_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') then IsNull(m.scmmeachul,0) else 0 end) as ten_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') then IsNull(m.scmmeachul,0) else 0 end) as ten_deliver_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then IsNull(m.scmmeachul,0) else 0 end) as ten_jungsan_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then m.ommeachul else 0 end) as ext_jungsan_sum " & vbCrLf
        strSql = strSql & " 		from " & vbCrLf
        strSql = strSql & " 		[db_statistics].[dbo].[tbl_extsite_orderMatching_master] m " & vbCrLf
        strSql = strSql & " 		where " & vbCrLf
        strSql = strSql & " 			1 = 1 " & vbCrLf
        strSql = strSql & " 			and ( " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 			) " & vbCrLf

        IF FRectOutMall <> "" then
            strSql = strSql & " 			AND m.sitename = '" & FRectOutMall & "' " & vbCrLf
        end IF

        strSql = strSql & " 		group by " & vbCrLf
        strSql = strSql & " 			m.scmactdate, m.sitename " & vbCrLf
        strSql = strSql & "  " & vbCrLf
        strSql = strSql & " 		union all " & vbCrLf
        strSql = strSql & "  " & vbCrLf
        strSql = strSql & " 		select " & vbCrLf
        strSql = strSql & " 			m.scmactdate as yyyymm, m.sitename " & vbCrLf
        strSql = strSql & " 			, isNull((select  sum(scmmeachul_d) from tbl_extsite_orderMatching_master tt where m.scmactdate = tt.scmactdate and m.sitename = tt.sitename),0) as payPrice " & vbCrLf
        strSql = strSql & " 			, '배송비' as itemtype " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_d else 0 end) as ten_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_d else 0 end) as ten_deliver_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_d else 0 end) as ten_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_d else 0 end) as ext_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_d else 0 end) as ten_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_d else 0 end) as ten_deliver_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_d else 0 end) as ten_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_d else 0 end) as ext_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_d else 0 end) as ten_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_d else 0 end) as ten_deliver_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_d else 0 end) as ten_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_d else 0 end) as ext_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_d else 0 end) as ten_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_d else 0 end) as ten_deliver_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_d else 0 end) as ten_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_d else 0 end) as ext_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_d else 0 end) as ten_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_d else 0 end) as ten_deliver_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_d else 0 end) as ten_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.ommeachul_d else 0 end) as ext_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') then m.scmmeachul_d else 0 end) as ten_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') then m.scmmeachul_d else 0 end) as ten_deliver_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then m.scmmeachul_d else 0 end) as ten_jungsan_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then m.ommeachul_d else 0 end) as ext_jungsan_sum " & vbCrLf
        strSql = strSql & " 		from " & vbCrLf
        strSql = strSql & " 		[db_statistics].[dbo].[tbl_extsite_orderMatching_master] m " & vbCrLf
        strSql = strSql & " 		where " & vbCrLf
        strSql = strSql & " 			1 = 1 " & vbCrLf
        strSql = strSql & " 			and ( " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf

        strSql = strSql & " 			) " & vbCrLf

        IF FRectOutMall <> "" then
            strSql = strSql & " 			AND m.sitename = '" & FRectOutMall & "' " & vbCrLf
        end IF

        strSql = strSql & " 		group by " & vbCrLf
        strSql = strSql & " 			m.scmactdate, m.sitename " & vbCrLf
        strSql = strSql & "  " & vbCrLf
        strSql = strSql & " 		union all " & vbCrLf
        strSql = strSql & "  " & vbCrLf
        strSql = strSql & " 		select " & vbCrLf
        strSql = strSql & " 			m.scmactdate as yyyymm, m.sitename " & vbCrLf
        strSql = strSql & " 			, isNull((select  sum(scmmeachul_m) from tbl_extsite_orderMatching_master tt where m.scmactdate = tt.scmactdate and m.sitename = tt.sitename),0) as payPrice " & vbCrLf
        strSql = strSql & " 			, '취소액' as itemtype " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_m else 0 end) as ten_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_m else 0 end) as ten_deliver_meachul_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_m else 0 end) as ten_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 0, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_m else 0 end) as ext_jungsan_yyyymm0 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_m else 0 end) as ten_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_m else 0 end) as ten_deliver_meachul_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_m else 0 end) as ten_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 1, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_m else 0 end) as ext_jungsan_yyyymm1 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_m else 0 end) as ten_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_m else 0 end) as ten_deliver_meachul_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_m else 0 end) as ten_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 2, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_m else 0 end) as ext_jungsan_yyyymm2 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmjsdate then m.scmmeachul_m else 0 end) as ten_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.scmdeliverdate then m.scmmeachul_m else 0 end) as ten_deliver_meachul_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.scmmeachul_m else 0 end) as ten_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when convert(varchar(7),DateAdd(MONTH, 3, ('" & minyyyymm & "' + '-01')), 121) = m.omjsdate then m.ommeachul_m else 0 end) as ext_jungsan_yyyymm3 " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_m else 0 end) as ten_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_m else 0 end) as ten_deliver_meachul_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.scmmeachul_m else 0 end) as ten_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) then m.ommeachul_m else 0 end) as ext_jungsan_null " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') then m.scmmeachul_m else 0 end) as ten_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.scmdeliverdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') then m.scmmeachul_m else 0 end) as ten_deliver_meachul_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then m.scmmeachul_m else 0 end) as ten_jungsan_sum " & vbCrLf
        strSql = strSql & " 			, sum(case when (m.omjsdate is NULL and (scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "')) or (omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') then m.ommeachul_m else 0 end) as ext_jungsan_sum " & vbCrLf
        strSql = strSql & " 		from " & vbCrLf
        strSql = strSql & " 		[db_statistics].[dbo].[tbl_extsite_orderMatching_master] m " & vbCrLf
        strSql = strSql & " 		where " & vbCrLf
        strSql = strSql & " 			1 = 1 " & vbCrLf
        strSql = strSql & " 			and ( " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmjsdate >= '" & minyyyymm & "' AND scmjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(omjsdate >= '" & minyyyymm & "' AND omjsdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf
        strSql = strSql & " 				or " & vbCrLf
        strSql = strSql & " 				( " & vbCrLf
        strSql = strSql & " 					(scmactdate >= '" & FRectStyyyymm & "' AND scmactdate <= '" & FRectEdyyyymm & "') " & vbCrLf
        strSql = strSql & " 					OR " & vbCrLf
        strSql = strSql & " 					(scmdeliverdate >= '" & minyyyymm & "' AND scmdeliverdate <= '" & maxyyyymm & "') " & vbCrLf
        strSql = strSql & " 				) " & vbCrLf

        strSql = strSql & " 			) " & vbCrLf

        IF FRectOutMall <> "" then
            strSql = strSql & " 			AND m.sitename = '" & FRectOutMall & "' " & vbCrLf
        end IF

        strSql = strSql & " 		group by " & vbCrLf
        strSql = strSql & " 			m.scmactdate, m.sitename " & vbCrLf
        strSql = strSql & " 	) T " & vbCrLf
        strSql = strSql & " order by " & vbCrLf
        strSql = strSql & " 	IsNull(T.yyyymm, '2999-01-01'), T.sitename " & vbCrLf
        strSql = strSql & " 	, (case " & vbCrLf
        strSql = strSql & " 			when T.itemtype = '상품' then 1 " & vbCrLf
        strSql = strSql & " 			when T.itemtype = '배송비' then 2 " & vbCrLf
        strSql = strSql & " 			when T.itemtype = '취소액' then 3 " & vbCrLf
        strSql = strSql & " 			else 100 end) " & vbCrLf

	    '// response.write strSql

        FCurrPage = 1
        FPageSize = 100

        rsSTSget.pagesize = FPageSize
        rsSTSget.CursorLocation = adUseClient
	    rsSTSget.Open strSql, dbSTSget, adOpenForwardOnly, adLockReadOnly

        FTotalCount = rsSTSget.RecordCount
        FResultCount = rsSTSget.RecordCount

        FTotalPage =  CLng(FTotalCount\FPageSize)
	    if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
		    FTotalPage = FtotalPage +1
	    end if
	    FResultCount = rsSTSget.RecordCount-(FPageSize*(FCurrPage-1))

        if FResultCount<1 then FResultCount=0

	    redim preserve FItemList(FResultCount)
	    i=0
	    if  not rsSTSget.EOF  then
		    rsSTSget.absolutepage = FCurrPage
		    do until rsSTSget.eof
			    set FItemList(i) = new CextJungsanMappingItem

			    FItemList(i).Fyyyymm						= rsSTSget("yyyymm")
                FItemList(i).Fsitename						= rsSTSget("sitename")
                FItemList(i).FpayPrice						= rsSTSget("payPrice")
                FItemList(i).Fitemtype						= rsSTSget("itemtype")

                FItemList(i).Ften_meachul_yyyymm0			= rsSTSget("ten_meachul_yyyymm0")
                FItemList(i).Ften_deliver_meachul_yyyymm0	= rsSTSget("ten_deliver_meachul_yyyymm0")
                FItemList(i).Ften_jungsan_yyyymm0			= rsSTSget("ten_jungsan_yyyymm0")
                FItemList(i).Fext_jungsan_yyyymm0			= rsSTSget("ext_jungsan_yyyymm0")

                FItemList(i).Ften_meachul_yyyymm1			= rsSTSget("ten_meachul_yyyymm1")
                FItemList(i).Ften_deliver_meachul_yyyymm1	= rsSTSget("ten_deliver_meachul_yyyymm1")
                FItemList(i).Ften_jungsan_yyyymm1			= rsSTSget("ten_jungsan_yyyymm1")
                FItemList(i).Fext_jungsan_yyyymm1			= rsSTSget("ext_jungsan_yyyymm1")

                FItemList(i).Ften_meachul_yyyymm2			= rsSTSget("ten_meachul_yyyymm2")
                FItemList(i).Ften_deliver_meachul_yyyymm2	= rsSTSget("ten_deliver_meachul_yyyymm2")
                FItemList(i).Ften_jungsan_yyyymm2			= rsSTSget("ten_jungsan_yyyymm2")
                FItemList(i).Fext_jungsan_yyyymm2			= rsSTSget("ext_jungsan_yyyymm2")

                FItemList(i).Ften_meachul_yyyymm3			= rsSTSget("ten_meachul_yyyymm3")
                FItemList(i).Ften_deliver_meachul_yyyymm3	= rsSTSget("ten_deliver_meachul_yyyymm3")
                FItemList(i).Ften_jungsan_yyyymm3			= rsSTSget("ten_jungsan_yyyymm3")
                FItemList(i).Fext_jungsan_yyyymm3			= rsSTSget("ext_jungsan_yyyymm3")

                FItemList(i).Ften_meachul_null				= rsSTSget("ten_meachul_null")
                FItemList(i).Ften_deliver_meachul_null		= rsSTSget("ten_deliver_meachul_null")
                FItemList(i).Ften_jungsan_null				= rsSTSget("ten_jungsan_null")
                FItemList(i).Fext_jungsan_null				= rsSTSget("ext_jungsan_null")

                FItemList(i).Ften_meachul_sum				= rsSTSget("ten_meachul_sum")
                FItemList(i).Ften_deliver_meachul_sum		= rsSTSget("ten_deliver_meachul_sum")
                FItemList(i).Ften_jungsan_sum				= rsSTSget("ten_jungsan_sum")
                FItemList(i).Fext_jungsan_sum				= rsSTSget("ext_jungsan_sum")

			    rsSTSget.moveNext
			    i=i+1
		    loop
	    end if
	    rsSTSget.Close
    End Function

 public Function fnGetextMatchingMaster
 dim strSql
 dim minscmdate, maxscmdate, minomdate, maxomdate, mindate, maxdate
 dim i
	strSql = " SELECT min(case when scmjsdate < omjsdate then scmjsdate else omjsdate end) as mindate, max(case when scmjsdate > omjsdate then scmjsdate else omjsdate end) as maxdate "
	strSql = strSql & " FROM db_statistics.dbo.tbl_extsite_orderMatching_master "
	strSql= strSql & " where scmactdate >='"&FRectStyyyymm&"' and scmactdate <='"&FRectEdyyyymm&"'"
	if FRectOutMall <> "" then
	strSql= strSql & " and sitename ='"&FRectOutMall&"'"
	end if
	''response.write strSql&"<br>"
	rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			mindate = rsSTSget("mindate")
			maxdate = rsSTSget("maxdate")

			if isNull(mindate) or isNull(maxdate) then
				FdiffDate = -1
			else
				FdiffDate = datediff("m",mindate,maxdate)
			end if
		else
			FdiffDate = -1
		END IF

		Fmindate = mindate
		Fmaxdate = maxdate
	rsSTSget.close

strSql = " select sitename, scmactdate "
strSql= strSql & " 	, isNull((select  sum(scmmeachul) from tbl_extsite_orderMatching_master  where scmactdate = tt.scmactdate and sitename = tt.sitename),0) as actitem "
strSql= strSql & "	, isNull((select  sum(scmmeachul_d) from tbl_extsite_orderMatching_master  where scmactdate = tt.scmactdate and sitename = tt.sitename),0) as actDe "
strSql= strSql & "	, isNull((select  sum(scmmeachul_m) from tbl_extsite_orderMatching_master  where scmactdate = tt.scmactdate and sitename = tt.sitename),0) as actM "
strSql = strSql & " , isNull(sum(scmitemN),0) as scmitemN "
strSql= strSql & "	, isNull(sum(scmDeN),0) as scmDeN"
strSql= strSql & "	, isNull(sum(scmMN),0) as scmMN"
strSql= strSql & "	, isNull(sum(somitemN),0) as somitemN "
strSql= strSql & "	, isNull(sum(somDeN),0) as somDeN "
strSql= strSql & "	, isNull(sum(somMN),0) as somMN "
for i =0 To FdiffDate
strSql= strSql & "	, isNull(sum(scmitem"&i&"),0) as scmitem"&i
strSql= strSql & "	, isNull(sum(scmDe"&i&"),0) as scmDe"&i
strSql= strSql & "	, isNull(sum(scmM"&i&"),0) as scmM"&i
strSql= strSql & "	, isNull(sum(somitem"&i&"),0) as somitem"&i
strSql= strSql & "	, isNull(sum(somDe"&i&"),0) as somDe"&i
strSql= strSql & "	, isNull(sum(somM"&i&"),0) as somM"&i
strSql= strSql & "	, isNull(sum(omitem"&i&"),0) as omitem"&i
strSql= strSql & "	, isNull(sum(omDe"&i&"),0) as omDe"&i
strSql= strSql & "	, isNull(sum(omM"&i&"),0) as omM"&i
next
strSql= strSql & " from ( "
strSql= strSql & " 		select     sitename, isNull(scmactdate,'NOMATCH') as scmactdate "
strSql = strSql & "		, case when jsdate is Null  then sum(scmitem) else 0 end as scmitemN"
strSql= strSql & "		, case when jsdate is Null then sum(scmDe) else 0 end as scmDeN"
strSql= strSql & "		, case when jsdate is Null then sum(scmM) else 0 end as scmMN"
strSql= strSql & "		, case when jsdate is Null then sum(somitem) else 0 end as somitemN"
strSql= strSql & "		, case when jsdate is Null then sum(somDe) else 0 end as somDeN"
strSql= strSql & "		, case when jsdate is Null then sum(somM) else 0 end as somMN"
for i =0 To FdiffDate
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(scmitem) else 0 end as scmitem"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(scmDe) else 0 end as scmDe"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(scmM) else 0 end as scmM"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(somitem) else 0 end as somitem"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(somDe) else 0 end as somDe"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(somM) else 0 end as somM"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(omitem) else 0 end as omitem"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(omDe) else 0 end as omDe"&i
strSql= strSql & "		, case when jsdate ='"&left(dateadd("m",i,mindate),7)&"' then sum(omM) else 0 end as omM"&i
next
strSql= strSql & "	 	from ( "
strSql= strSql & "		select sitename, scmactdate "
strSql= strSql & "			, scmjsdate as jsdate "
'strSql= strSql & "			, case when (scmactdate is   null or scmjsdate is null) then sum(scmmeachul) else sum(scmmeachul+scmmeachul_m)  end as scmitem"
strSql= strSql & "			, sum(scmmeachul) as scmitem"
strSql= strSql & "			, sum(scmmeachul_D) as scmDe"
'strSql= strSql & "			, case when (scmactdate is   null or scmjsdate is null) then sum(scmmeachul_m) else 0  end as scmM "
strSql= strSql & "			,  sum(scmmeachul_m)  as scmM "
strSql= strSql & "			, 0 as somitem , 0 as somDe, 0 as somM"
strSql= strSql & "			, 0 as omitem , 0 as omDe, 0 as omM "
strSql= strSql & "		from db_statistics.dbo.tbl_extsite_orderMatching_master  "
strSql= strSql & "		where ( ( scmactdate>='"&FRectStyyyymm&"' and scmactdate<='"&FRectEdyyyymm&"' ) or (scmjsdate>='"&mindate&"' and scmjsdate<='"&maxdate&"') ) "
IF FRectOutMall <> "" then
strSql= strSql & "			and sitename ='"&FRectOutMall&"'"
end IF
strSql= strSql & "		group by sitename, scmactdate, scmjsdate "
strSql= strSql & "		union all "
strSql= strSql & "		select sitename, scmactdate "
strSql= strSql & "		,  omjsdate as jsdate "
strSql= strSql & "		, 0 as scmitem , 0 as scmDe, 0 as scmM "
'strSql= strSql & "		, case when (scmactdate is   null or omjsdate is null) then sum(scmmeachul) else sum(scmmeachul+scmmeachul_m)  end as somitem "
strSql= strSql & "		, sum(scmmeachul) as somitem "
strSql= strSql & "		, sum(scmmeachul_D) as somDe "
'strSql= strSql & "		, case when (scmactdate is   null or omjsdate is null) then sum(scmmeachul_m) else 0  end as somM "
strSql= strSql & "		, sum(scmmeachul_m)   as somM "
'strSql= strSql & "		, case when (scmactdate is   null or omjsdate is null) then sum(ommeachul) else sum(ommeachul+ommeachul_m)  end as omitem "
strSql= strSql & "		,   sum(ommeachul)  as omitem "
strSql= strSql & "		, sum(ommeachul_D) as omDe "
'strSql= strSql & "		, case when (scmactdate is   null or omjsdate is null) then sum(ommeachul_m) else 0  end as omM "
strSql= strSql & "		,  sum(ommeachul_m)   as omM "
strSql= strSql & "		from db_statistics.dbo.tbl_extsite_orderMatching_master "
strSql= strSql & "		where ( ( scmactdate>='"&FRectStyyyymm&"' and scmactdate<='"&FRectEdyyyymm&"' ) or (omjsdate>='"&mindate&"' and omjsdate<='"&maxdate&"') ) "
IF FRectOutMall <> "" then
strSql= strSql & "			and sitename ='"&FRectOutMall&"'"
end IF
strSql= strSql & "		group by sitename, scmactdate, omjsdate "
strSql= strSql & "	) as T "
strSql= strSql & " group by sitename, scmactdate, jsdate "
strSql= strSql & ") as tt "
strSql= strSql & "group by sitename, scmactdate "
strSql= strSql & " order by   isNull(scmactdate,'NOMATCH'), sitename "
''response.write strSql&"<br>"
rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			fnGetextMatchingMaster	= rsSTSget.getRows()
		END IF
	rsSTSget.close
 End Function


 public Function fnGetextMatchingData
 dim strSql
 dim minscmdate, maxscmdate, minomdate, maxomdate, mindate, maxdate
 dim i
	strSql = " SELECT min(scmjsdate) as minscmdate, max(scmjsdate) as maxscmdate, min(omjsdate) as minomdate, max(omjsdate) as maxomdate "
	strSql = strSql & " FROM db_statistics.dbo.tbl_extsite_orderMatching "
	strSql= strSql & " where scmactdate ='"&FRectyyyymm&"' and sitename ='"&FRectOutMall&"'"
	'response.write strSql&"<br>"
	rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			minscmdate = rsSTSget("minscmdate")
			maxscmdate = rsSTSget("maxscmdate")
			minomdate = rsSTSget("minomdate")
			maxomdate	 = rsSTSget("maxomdate")

			mindate = minscmdate
			if minscmdate > minomdate then mindate = minomdate
			maxdate = maxscmdate
			if maxscmdate < maxomdate then maxdate = maxomdate
			if isNull(mindate) or isNull(maxdate) then
				FdiffDate = -1
			else
				FdiffDate = datediff("m",mindate,maxdate)
			end if
		else
			FdiffDate = -1
		END IF

		Fmindate = mindate
		Fmaxdate = maxdate
	rsSTSget.close


	strSql = " SELECT isNull(scmactdate,'미매칭') "
	strSql = strSql & " ,  isNull((select isNull(sum(scmmeachul),0) from db_statistics.dbo.tbl_extSite_orderMatching  where sitename = '"&FRectOutMall&"' and scmactdate = T.scmactdate),0) as actmeachul "
	strSql = strSql & " ,  isNull((select isNull(sum(scmmeachul_d),0) from db_statistics.dbo.tbl_extSite_orderMatching  where sitename = '"&FRectOutMall&"' and scmactdate = T.scmactdate),0)  as actmeachul_d "
	strSql = strSql & " ,isNull(sum(scmN),0), isNull(sum(scmN_d),0), isNull(sum(omN),0), isNull(sum(omN_d),0) "
	for i =0 To FdiffDate
	strSql = strSql & ", isNull(sum(scm"&i&"),0), isNull(sum(scm"&i&"_d),0)"
	next
	for i =0 To FdiffDate
	strSql = strSql & ", isNull(sum(om"&i&"),0), isNull(sum(om"&i&"_d),0)"
	next
	for i =0 To FdiffDate
	strSql = strSql & ", isNull(sum(omscm"&i&"),0), isNull(sum(omscm"&i&"_d),0)"
	next
	strSql= strSql & " , isNull(sum(omscmN),0), isNull(sum(omscmN_d),0)"
	strSql = strSql & " , isNull(sum(scmCancelN),0), isNull(sum(omCancelN),0) "
	strSql = strSql & " , isNull((select sum(ommeachul+ommeachul_d) from db_statistics.dbo.tbl_extSite_orderMatching  where sitename = '"&FRectOutMall&"' and scmactdate is null and omjsdate ='"&FRectyyyymm&"' and isMYN='Y'),0) as actCancelN "
	strSql = strSql & " FROM ( "
	strSql = strSql & "		SELECT scmactdate "
	strSql = strSql & "		,'0' as actmeachul"
	strSql = strSql & " 	, '0' as actmeachul_d "
	for i =0 To FdiffDate
	strSql = strSql & "		, case when scmjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null) then sum(scmmeachul) else 0 end as scm"&i
	strSql = strSql & "		, case when scmjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null)  then sum(scmmeachul_d) else 0 end as scm"&i&"_d "
	strSql = strSql & "		,'0' as om"&i&",'0' as om"&i&"_d,'0' as omscm"&i&",'0' as omscm"&i&"_d"
	next
	strSql = strSql & "		, case when scmjsdate is null and isMYN ='N' then sum(scmmeachul) else 0 end as scmN "
	strSql = strSql & "		, case when scmjsdate is null and isMYN ='N' then sum(scmmeachul_d) else 0 end as scmN_d "
	strSql = strSql & "		,'0' as omN, '0' as omN_d ,'0' as omscmN, '0' as omscmN_d "
	strSql = strSql & "		, case when  scmjsdate is null and isMYN ='Y' then   sum(scmmeachul+scmmeachul_d) else 0 end as scmCancelN "
	strSql = strSql & "		,'0' as omCancelN "
	strSql = strSql & "		FROM db_statistics.dbo.tbl_extSite_orderMatching as A"
	strSql = strSql & "		WHERE sitename ='"&FRectOutMall&"' "
	strSql = strSql & "			and ((scmjsdate >= '"&mindate&"' and scmjsdate <='"&maxdate&"') or (scmactdate ='"&FRectyyyymm&"' and scmjsdate is null ))"
	strSql = strSql & "		GROUP BY scmactdate, scmjsdate,isMYN "
	strSql = strSql & "		union all "
	strSql = strSql & "		SELECT scmactdate, '0' as actmeachul, '0' as actmeachul_d "
	for i =0 To FdiffDate
	strSql = strSql & "		, '0' as scm"&i&", '0' as scm"&i&"_d"
	strSql = strSql & "		, case when omjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null)  then sum(ommeachul) else 0 end as om"&i
	strSql = strSql & "		, case when omjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null)  then sum(ommeachul_d) else 0 end as om"&i&"_d "
	strSql = strSql & "		, case when omjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null)  then sum(scmmeachul) else 0 end as omscm"&i
	strSql = strSql & "		, case when omjsdate =convert(varchar(7),'"&dateadd("m",i,mindate)&"',121)  and ((scmactdate is null and isMYN ='N') or scmactdate is not null)  then sum(scmmeachul_d) else 0 end as omscm"&i&"_d "
	next
	strSql = strSql & "		, '0' as scmN, '0' as scmN_d "
	strSql = strSql & "		, case when omjsdate is null  and isMYN ='N' then sum(ommeachul) else 0 end as omN "
	strSql = strSql & "		, case when omjsdate is null  and isMYN ='N' then sum(ommeachul_d) else 0 end as omN_d "
	strSql = strSql & "		, case when omjsdate is null  and isMYN ='N' then sum(scmmeachul) else 0 end as omscmN "
	strSql = strSql & "		, case when omjsdate is null  and isMYN ='N' then sum(scmmeachul_d) else 0 end as omscmN_d "
	strSql = strSql & "		, '0' as scmCancelN "
	strSql= strSql & "		,  case when omjsdate is null and isMYN ='Y' then sum(scmmeachul+scmmeachul_d) else 0 end as omCancelN  "
	strSql = strSql & "		FROM db_statistics.dbo.tbl_extSite_orderMatching "
	strSql = strSql & "		WHERE sitename ='"&FRectOutMall&"' "
	strSql = strSql & "			and ((omjsdate >= '"&mindate&"' and omjsdate <='"&maxdate&"') or (scmactdate ='"&FRectyyyymm&"' and omjsdate is null ))"
	strSql = strSql & "		GROUP BY scmactdate, omjsdate , isMYN "
	strSql = strSql & " ) AS T"
	strSql = strSql & "	GROUP BY scmactdate  "
	strSql = strSql & " ORDER BY isNull(scmactdate,'미매칭') "
'response.write strSql
'response.end

	rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			fnGetextMatchingData	= rsSTSget.getRows()
		END IF
	rsSTSget.close
 End Function


public FRectscmJsDate
public FRectscmDeliverDate
public FRectomJsDate
public FRectItemType
public FRectDiffYN
public FRectItemDiv
public FRectIsMYN

public FRectJsType
public FRectJsDate

public FPSize
public FCPage
public FTotCnt
public FRectSort
public FRectOrderserial

 public Function fnGetextMatchingItem
  Dim strSql
  dim iSPageNo
  iSPageNo =  FPSize*(FCPage-1)
  if FRectyyyymm = "" or FRectyyyymm ="NOMATCH" THEN FRectyyyymm = "N"

  		strSql = "db_statistics.[dbo].[usp_Ten_extsite_orderItemMatching_getCount]('"&FRectOutMall&"','"&FRectyyyymm&"','"&FRectscmJsDate&"','"&FRectomJsDate&"','"&FRectItemDiv&"','"&FRectOrderserial&"','"&FRectIsMYN&"','"&FRectscmDeliverDate&"')"
		rsSTSget.Open strSql,dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF not rsSTSget.EOF THEN
			FTotCnt = rsSTSget(0)
		End if
		rsSTSget.close

		strSql = "db_statistics.[dbo].[usp_Ten_extsite_orderItemMatching_getData]('"&FRectOutMall&"','"&FRectyyyymm&"','"&FRectscmJsDate&"','"&FRectomJsDate&"','"&FRectItemDiv&"','"&FRectIsMYN&"','"&FRectOrderserial&"','"&FRectSort&"','"&iSPageNo&"','"&FPSize&"','"&FRectscmDeliverDate&"')"
		response.write strSql
		rsSTSget.Open strSql,dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF not rsSTSget.EOF THEN
			fnGetextMatchingItem = rsSTSget.getRows()
		End if
		rsSTSget.close
 End Function

 public Fscmitemno
 public Fomitemno
 public Fscmsellprice
 public Fscmmeachul
 public Fscmbuycash
 public Fomsellprice
 public Fommeachul
 public Fombuycash
 public FextTenCouponPrice
 public FextOwnCouponPrice
 public FreducedPrice
 public FallAtDiscount

 public Function fnGetextMatchingItemSUM
  Dim strSql
		strSql = "db_statistics.[dbo].[usp_Ten_extsite_orderItemMatching_getSUM]('"&FRectOutMall&"','"&FRectyyyymm&"','"&FRectscmJsDate&"','"&FRectomJsDate&"','"&FRectItemDiv&"','"&FRectOrderserial&"','"&FRectIsMYN&"','"&FRectscmDeliverDate&"')"
		rsSTSget.Open strSql,dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF not rsSTSget.EOF THEN
			 Fscmitemno = rsSTSget("scmitemno")
			 Fomitemno = rsSTSget("omitemno")
			 Fscmsellprice = rsSTSget("scmsellprice")
			 Fscmmeachul = rsSTSget("scmmeachul")
			 Fscmbuycash = rsSTSget("scmbuycash")
			 Fomsellprice = rsSTSget("omsellprice")
			 Fommeachul = rsSTSget("ommeachul")
			 Fombuycash = rsSTSget("ombuycash")
			 FextTenCouponPrice = rsSTSget("extTenCouponPrice")
			 FextOwnCouponPrice = rsSTSget("extOwnCouponPrice")
			 FreducedPrice = rsSTSget("reducedPrice")
			FallAtDiscount = rsSTSget("allAtDiscount")

		End if
	rsSTSget.close
 End Function

public Function fnGetextMappingItem
 dim strSql
 strSql ="select oi.orderserial, oi.itemid, oi.itemoption, scmitemno, omitemno, isNull(scmmeachul,0), isNull(ommeachul,0), isNull(scmsellprice,0), isNull(omsellprice,0) , scmjsdate, omjsdate "
 strSql =strSql & " ,d.makerid, isNull(orgitemcost,0), isNull(itemcost,0), isNull(reducedprice,0), isNull(buycash,0), isNull(upchejungsancash,0), isNull(extitemcost,0)"
 strSql = strSql & " , isNull(extreducedprice,0), isNull(extowncouponprice,0), isNull(exttencouponprice,0), isNull(omcommprice,0), isNull(omjungsanprice,0) "
 strSql= strSql & " from db_statistics.dbo.tbl_extsite_orderitemMatching as oi"
 strSql= strSql & "	left outer join db_statistics.dbo.tbl_order_detail_log as d "
 strSql= strSql & "	on oi.orderserial = d.orderserial "
 strSql= strSql & "	and oi.itemid = d.itemid "
 strSql= strSql & "	and oi.itemoption = d.itemoption "
 strSql= strSql & " and oi.suborderserial = d.suborderserial "
 strSql= strSql & " left outer join db_statistics.dbo.tbl_xsite_jungsandata as j "
 strSql= strSql & "	on oi.orderserial = isNull(j.orgorderserial,'NOMATCH') "
 strSql= strSql & "	and oi.itemid = j.itemid "
 strSql= strSql & "	and oi.itemoption = j.itemoption "
 strSql= strSql & "	and oi.extorderserial = j.extOrderserial "
 strSql= strSql & "	and oi.extorderserseq = j.extOrderserSeq  "
 strSql = strSql & " where "
 if FRectscmJsDate ="" then
 strSql = strSql & "  scmjsdate is null  "
else
	strSql = strSql & " convert(varchar(7),scmjsdate,121) ='"&FRectscmJsDate&"' "
end if
if FRectomJsDate = "" then
	strSql = strSql & " and omjsdate is null "
else
 strSql = strSql & " and convert(varchar(7),omjsdate,121) ='"&FRectomJsDate&"' "
end if
if FRectItemType ="D" then
	strSql = strSql & " and oi.itemid = 0  "
else
	strSql = strSql & " and (oi.itemid <> 0 or oi.itemid is null) "
end if
if FRectDiffYN ="Y" then
    strSql = strSql & " and ( oi.scmitemno <> oi.omitemno or oi.scmmeachul <> oi.ommeachul)  "
end if
 strSql = strSql & " and sitename ='"&FRectOutMall&"'"
 strSql = strSql & " order by oi.orderserial ,oi.itemid ,oi.itemoption"
' response.write strSql

 rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			fnGetextMappingItem = rsSTSget.getRows()
		End if
	rsSTSget.close
end Function
End Class
Class CextJungsanDiff
public FPSize
public FCPage
public FTotCnt

public FCGFDate
public FCGTDate
public FCFFDate
public FCFTDate
public FSellsite
public FRectST
public FRectErr

public FextMeachul
public FlogMeachul



public Function fnGetextJsDiffList
dim strSql, strSqlAdd
dim iSPageNo
	iSPageNo =  FPSize*(FCPage-1)
 strSqlAdd = ""
 if FRectErr ="Y" then
 	strSqlAdd = strSqlAdd &" and o.orderserial is null "
end if

 strSql = " select count(j.orgorderserial)  "
 strSql = strSql & " from db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  left outer join "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' and  d.beasongdate>='"&FCGFDate&"' and d.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  where j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & strSqlAdd
 rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			FTotCnt = rsSTSget(0)
		End if
	rsSTSget.close


 strSql = " select  sum(extitemno*exttenmeachulprice) as extMeachul "
 strSql = strSql & " ,  sum(o.itemno*o.itemcost)  as logMeachul "
 strSql = strSql & " from db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  left outer join "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' and  d.beasongdate>='"&FCGFDate&"' and d.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  where j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & strSqlAdd
 rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			FextMeachul = rsSTSget("extMeachul")
			FlogMeachul = rsSTSget("logMeachul")
		End if
	rsSTSget.close

 strSql = " select j.sellsite, j.orgorderserial, j.itemid, j.itemoption, j.extorderserial, extorderserseq, extitemno, exttenmeachulprice , extcommprice  "
 strSql= strSql & " , o.orderserial, o.itemid , o.itemoption , o.itemno, o.itemcost "
 strSql = strSql & " from db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  left outer join "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' and  d.beasongdate>='"&FCGFDate&"' and d.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  where j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & strSqlAdd
 strSql = strSql & " order by j.orgorderserial desc"
 strSql = strSql & "    offset "&iSPageNo&" rows "
	strSql = strSql & "  fetch next "&FPSize&"  rows only   "
	'response.write strSql
	rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			fnGetextJsDiffList = rsSTSget.getRows()
		End if
	rsSTSget.close
end Function



public Function fnGetlogJsDiffList
dim strSql, strSqlAdd
dim iSPageNo
	iSPageNo =  FPSize*(FCPage-1)
 strSqlAdd = ""
 if FRectErr ="Y" then
 	strSqlAdd = strSqlAdd &" and j.orgorderserial is null "
end if

 strSql = " select count(o.orderserial)  "
 strSql = strSql & " from "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o "
 strSql = strSql & "  left outer join db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  and j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & "  where  o.beasongdate>='"&FCGFDate&"' and o.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & strSqlAdd
 rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			FTotCnt = rsSTSget(0)
		End if
	rsSTSget.close


 strSql = " select  sum(extitemno*exttenmeachulprice) as extMeachul "
 strSql = strSql & " ,  sum(o.itemno*o.itemcost)  as logMeachul "
  strSql = strSql & " from "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o "
 strSql = strSql & "  left outer join db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  and j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & "  where  o.beasongdate>='"&FCGFDate&"' and o.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & strSqlAdd
 rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			FextMeachul = rsSTSget("extMeachul")
			FlogMeachul = rsSTSget("logMeachul")
		End if
	rsSTSget.close

 strSql = " select j.sellsite, j.orgorderserial, j.itemid, j.itemoption, j.extorderserial, extorderserseq, extitemno, exttenmeachulprice , extcommprice  "
 strSql= strSql & " , o.orderserial, o.itemid , o.itemoption , o.itemno, o.itemcost "
 strSql = strSql & " from "
 strSql = strSql & "  (select m.sitename, m.targetGbn,d.* "
 strSql = strSql & " 	from db_statistics.dbo.tbl_order_master_log as m  "
 strSql = strSql & " 	inner join db_statistics.dbo.tbl_order_detail_log as d "
 strSql = strSql & " 	with (index(IX_tbl_order_detail_log_beasongdate)) "
 strSql = strSql & " 	on m.orderserial = d.orderserial and m.suborderserial = d.suborderserial "
 strSql = strSql & " 	where IsNull(m.targetGbn, 'ON') = 'ON' "
 strSql = strSql & " 			 and m.sitename = '"&FSellsite&"'   "
 strSql = strSql & "  ) as o "
 strSql = strSql & "  left outer join db_statistics.dbo.tbl_xSite_JungsanData as j "
 strSql = strSql & "  on o.orderserial = j.orgorderserial and o.sitename = j.sellsite and o.itemid = j.itemid and  o.itemoption = j.itemoption "
 strSql = strSql & "  and j.extMeachulDate >='"&FCFFDate&"' and j.extMeachulDate < '"&FCFTDate&"' and j.sellsite ='"&FSellsite&"' "
 strSql = strSql & "  where  o.beasongdate>='"&FCGFDate&"' and o.beasongdate<'"&FCGTDate&"' "
 strSql = strSql & strSqlAdd
 strSql = strSql & " order by o.orderserial desc"
 strSql = strSql & "    offset "&iSPageNo&" rows "
	strSql = strSql & "  fetch next "&FPSize&"  rows only   "
	'response.write strSql
	rsSTSget.Open strSql,dbSTSget
		IF not rsSTSget.EOF THEN
			fnGetlogJsDiffList = rsSTSget.getRows()
		End if
	rsSTSget.close
end Function
End Class
%>
