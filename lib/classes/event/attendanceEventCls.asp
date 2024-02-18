<%
Class AttendanceEventCls

	Public Fidx
	Public Flink_evt_code
	Public Fmain_image
	Public Fmain_image_link
	Public Fbg_color
	Public Fbutton_before_day_color
	Public Fbutton_before_point_color
	Public Fbutton_before_bg_color
	Public Fbutton_after_day_color
	Public Fbutton_after_point_color
	Public Fbutton_after_bg_color
	Public Fcheck_area_bg_color
	Public Fcheck_title_color
	Public Fcheck_button_bg_color
	Public Fcheck_button_title_color
	Public Fcheck_etc_contents
	Public Fcheck_etc_contents_color
	Public Falarm_bg_color
	Public Falarm_etc_contents
	Public Fdeeplink
	Public Fevt_name
    Public Fmo_main_image
    Public Fmo_main_image2
    Public Fbutton_today_ring_color
    Public Fpopup_bubble_bg_color
    Public Fpopup_bubble_text_color
	Public Fmileage_summary

    Public FrectEvt_Code
	
	Private Sub Class_Initialize()

	End Sub
	Private Sub Class_Terminate()

	End Sub

	'// one brand
	public Sub getOneContents()
        dim sqlStr
       	sqlStr = "SELECT a.idx , a.link_evt_code, a.main_image, a.main_image_link, a.bg_color , a.button_before_day_color, a.button_before_point_color,"
		sqlStr = sqlStr & " a.button_before_bg_color, a.button_after_day_color, a.button_after_point_color, a.button_after_bg_color, a.check_area_bg_color,"
		sqlStr = sqlStr & " a.check_title_color, a.check_button_bg_color, a.check_button_title_color, a.check_etc_contents, a.check_etc_contents_color,"
        sqlStr = sqlStr & " a.alarm_bg_color, a.alarm_etc_contents, a.deeplink, e.evt_name, a.mo_main_image, a.mo_main_image2, a.button_today_ring_color,"
        sqlStr = sqlStr & " a.popup_bubble_bg_color, a.popup_bubble_text_color, a.mileage_summary"
        sqlStr = sqlStr & " FROM [db_event].[dbo].[tbl_event_attendance] AS a WITH(NOLOCK)"
		sqlStr = sqlStr & " right JOIN [db_event].[dbo].[tbl_event] AS e WITH(NOLOCK)"
		sqlStr = sqlStr & " ON a.evt_code = e.evt_code"
        sqlStr = sqlStr & " WHERE a.evt_code=" & CStr(FrectEvt_Code)
        rsget.Open SqlStr, dbget, 1
        if Not rsget.Eof then
			Fidx = rsget("idx")
            Flink_evt_code = rsget("link_evt_code")
			Fmain_image = rsget("main_image")
			Fmain_image_link = rsget("main_image_link")
			Fbg_color = rsget("bg_color")
			Fbutton_before_day_color = rsget("button_before_day_color")
			Fbutton_before_point_color = rsget("button_before_point_color")
			Fbutton_before_bg_color = rsget("button_before_bg_color")
			Fbutton_after_day_color = rsget("button_after_day_color")
			Fbutton_after_point_color = rsget("button_after_point_color")
			Fbutton_after_bg_color = rsget("button_after_bg_color")
			Fcheck_area_bg_color = rsget("check_area_bg_color")
			Fcheck_title_color	= rsget("check_title_color")
			Fcheck_button_bg_color = rsget("check_button_bg_color")
			Fcheck_button_title_color = rsget("check_button_title_color")
			Fcheck_etc_contents = rsget("check_etc_contents")
            Fcheck_etc_contents_color = rsget("check_etc_contents_color")
            Falarm_bg_color = rsget("alarm_bg_color")
            Falarm_etc_contents = rsget("alarm_etc_contents")
			Fdeeplink = rsget("deeplink")
            Fevt_name = rsget("evt_name")
            Fmo_main_image = rsget("mo_main_image")
            Fmo_main_image2 = rsget("mo_main_image2")
            Fbutton_today_ring_color = rsget("button_today_ring_color")
            Fpopup_bubble_bg_color = rsget("popup_bubble_bg_color")
            Fpopup_bubble_text_color = rsget("popup_bubble_text_color")
			Fmileage_summary = rsget("mileage_summary")
        end if
        rsget.Close
    end Sub

End Class
%>