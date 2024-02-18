<%
''<option value="00" >배송문의</option>
''<option value="01" >주문문의</option>
''<option value="02" >상품문의</option>
''<option value="03" >재고문의</option>
''<option value="04" >취소문의</option>
''<option value="05" >환불문의</option>
''<option value="06" >교환문의</option>
''<option value="07" >AS문의</option>    
''<option value="08" >이벤트문의</option>
''<option value="09" >증빙서류문의</option>    
''<option value="10" >시스템문의</option>
''<option value="11" >회원제도문의</option>
''<option value="12" >회원정보문의</option>
''<option value="13" >당첨문의</option>
''<option value="14" >반품문의</option>
''<option value="15" >입금문의</option>
''<option value="16" >오프라인문의</option>
''<option value="17" >쿠폰/마일리지문의</option>
''<option value="18" >결제방법문의</option>
''<option value="20" >기타문의</option>

class CReportMasterItemList
    public Fqadiv
 	public Fcount
    public FqadivName
    
    public function GetQadivName()
        if Fqadiv="00" then
            GetQadivName = "배송문의"
        elseif Fqadiv="01" then
            GetQadivName = "주문문의"
        elseif Fqadiv="02" then
            GetQadivName = "상품문의"
        elseif Fqadiv="03" then
            GetQadivName = "재고문의"
        elseif Fqadiv="04" then
            GetQadivName = "취소문의"
        elseif Fqadiv="05" then
            GetQadivName = "환불문의"
        elseif Fqadiv="06" then
            GetQadivName = "교환문의"
        elseif Fqadiv="07" then
            GetQadivName = "As문의"
        elseif Fqadiv="08" then
            GetQadivName = "이벤트문의"
        elseif Fqadiv="09" then
            GetQadivName = "증빙서류문의"
        elseif Fqadiv="10" then
            GetQadivName = "시스템문의"
        elseif Fqadiv="11" then
            GetQadivName = "회원제도문의"
        elseif Fqadiv="12" then
            GetQadivName = "개인정보관련"
        elseif Fqadiv="13" then
            GetQadivName = "당첨문의"
        elseif Fqadiv="14" then
            GetQadivName = "반품문의"
        elseif Fqadiv="15" then
            GetQadivName = "입금문의"
        elseif Fqadiv="16" then
            GetQadivName = "오프라인문의"    
        elseif Fqadiv="17" then
            GetQadivName = "쿠폰/마일리지문의"    
        elseif Fqadiv="18" then
            GetQadivName = "결제방법문의"    
            
        elseif Fqadiv="20" then
            GetQadivName = "기타문의"
        else
            GetQadivName = Fqadiv
        end if
    end function

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end class

class CReportMaster
	public FMasterItemList()
	public FRectStart
	public FRectEnd
	public FResultCount
	public Ftotalcount

	Private Sub Class_Initialize()
		'redim preserve FMasterItemList(0)
		redim  FMasterItemList(0)

		FResultCount = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub

	public Sub SearchReport()

	    dim sql,i
'		sql = "select count(qadiv) as count from [db_cs].[10x10].tbl_myqna" + vbcrlf
'		sql = sql + " where regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
'		sql = sql + " and regdate < '" + Cstr(FRectEnd) + "'"
'
'		rsget.Open sql,dbget,1
'
'		if  not rsget.EOF  then
'			Ftotalcount = rsget("count")
'		end if
'		rsget.close

		sql = "select c.qadiv, c.qadivname, count(q.id) as count "
		sql = sql + " from [db_cs].[10x10].tbl_myqna q" + vbcrlf
		sql = sql + "   left join [db_cs].[dbo].tbl_myqna_comm_code c"
		sql = sql + "   on q.qadiv=c.qadiv"
		sql = sql + " where q.regdate >= '" + Cstr(FRectStart) + "'" + vbcrlf
		sql = sql + " and q.regdate < '" + Cstr(FRectEnd) + "'" + vbcrlf
		sql = sql + " and q.isusing='Y'"
		sql = sql + " and q.dispyn='Y'"
		sql = sql + " group by all c.qadiv, c.qadivname" + vbcrlf
		sql = sql + " order by c.qadiv asc"

		rsget.Open sql,dbget,1

		FResultCount = rsget.recordcount
		redim preserve FMasterItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
		    do until rsget.EOF
				set FMasterItemList(i) = new CReportMasterItemList
				FMasterItemList(i).Fqadiv = rsget("qadiv")
				FMasterItemList(i).Fcount = rsget("count")
				FMasterItemList(i).FqadivName = rsget("qadivname")
				rsget.movenext
				i=i+1
		    loop
		end if
		rsget.close

	end sub

end class

%>
