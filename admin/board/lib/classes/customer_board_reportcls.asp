<%
''<option value="00" >��۹���</option>
''<option value="01" >�ֹ�����</option>
''<option value="02" >��ǰ����</option>
''<option value="03" >�����</option>
''<option value="04" >��ҹ���</option>
''<option value="05" >ȯ�ҹ���</option>
''<option value="06" >��ȯ����</option>
''<option value="07" >AS����</option>    
''<option value="08" >�̺�Ʈ����</option>
''<option value="09" >������������</option>    
''<option value="10" >�ý��۹���</option>
''<option value="11" >ȸ����������</option>
''<option value="12" >ȸ����������</option>
''<option value="13" >��÷����</option>
''<option value="14" >��ǰ����</option>
''<option value="15" >�Աݹ���</option>
''<option value="16" >�������ι���</option>
''<option value="17" >����/���ϸ�������</option>
''<option value="18" >�����������</option>
''<option value="20" >��Ÿ����</option>

class CReportMasterItemList
    public Fqadiv
 	public Fcount
    public FqadivName
    
    public function GetQadivName()
        if Fqadiv="00" then
            GetQadivName = "��۹���"
        elseif Fqadiv="01" then
            GetQadivName = "�ֹ�����"
        elseif Fqadiv="02" then
            GetQadivName = "��ǰ����"
        elseif Fqadiv="03" then
            GetQadivName = "�����"
        elseif Fqadiv="04" then
            GetQadivName = "��ҹ���"
        elseif Fqadiv="05" then
            GetQadivName = "ȯ�ҹ���"
        elseif Fqadiv="06" then
            GetQadivName = "��ȯ����"
        elseif Fqadiv="07" then
            GetQadivName = "As����"
        elseif Fqadiv="08" then
            GetQadivName = "�̺�Ʈ����"
        elseif Fqadiv="09" then
            GetQadivName = "������������"
        elseif Fqadiv="10" then
            GetQadivName = "�ý��۹���"
        elseif Fqadiv="11" then
            GetQadivName = "ȸ����������"
        elseif Fqadiv="12" then
            GetQadivName = "������������"
        elseif Fqadiv="13" then
            GetQadivName = "��÷����"
        elseif Fqadiv="14" then
            GetQadivName = "��ǰ����"
        elseif Fqadiv="15" then
            GetQadivName = "�Աݹ���"
        elseif Fqadiv="16" then
            GetQadivName = "�������ι���"    
        elseif Fqadiv="17" then
            GetQadivName = "����/���ϸ�������"    
        elseif Fqadiv="18" then
            GetQadivName = "�����������"    
            
        elseif Fqadiv="20" then
            GetQadivName = "��Ÿ����"
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
