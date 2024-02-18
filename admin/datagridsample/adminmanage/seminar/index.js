$(function () {
    var roomInfoList = [];
    var roomInfo = new DevExpress.data.DataSource({
        key : "id",
        load: function () {
            var d = $.Deferred();
            $.getJSON('roominfo.asp').done(function (data) {
                d.resolve(data); 
                roomInfoList = data;
            }).fail(d.reject);
            return d.promise();
        }
    });

    $("#scheduler").dxScheduler({
        height: 600,
        dataSource: "roomlist.asp",
        showAllDayPanel: false,
        shadeUntilCurrentTime: true,
        indicatorUpdateInterval: 30000, // 30마다 자동 갱신
        views: [{
            type : "month",
            name : "Month",
        }, {
            type : "day",
            name : "Day-room",
            groups : ["roomId"],
            startDayHour : 9,
            endDayHour : 24,
        }],
        currentView: "day",
        height: 680,
        currentDate: new Date(),
        resources: [{
            fieldExpr: "roomId",
            allowMultiple: false,
            dataSource: roomInfo,
            label: "Room"
        }],
        editing : {
            allowDeleting : false,
            allowAdding : false,
            allowUpdating : false,
        },
        onContentReady(e) {
            e.component.scrollTo(new Date());
        },
        appointmentTooltipTemplate: function(data) {
            return getAppointmentTemplate(data);
        },
    });

    function getAppointmentTemplate(data) {
        var backgroundColorStyle = data.roomId ? "style='background-color:" + getAppointmentColor(data.roomId) + ";'" : "";
        var markup = $("<div class='appointment-content'>" +
                    "<div class='appointment-badge'" + backgroundColorStyle + ">" + getAppointmentRoomName(data.roomId).toString()[0] + "</div>" +
                    "<div class='appointment-text'>" + data.text + "</div>" + 
                    "<div class='appointment-text'>회의실 : " + getAppointmentRoomName(data.roomId) + "</div>" + 
                    "<div class='appointment-text'>등록자 : " + data.username + "</div>" + 
                    "<div class='appointment-dates'>" + Globalize.formatDate(new Date(data.startDate), { skeleton: "MMMd" }) + 
                        " , " + Globalize.formatDate(new Date(data.startDate), { time: "short" }) +
                            " - " + Globalize.formatDate(new Date(data.endDate), { time: "short" }) +
                    "</div>" + 
                    "</div>" + 
                    "</div>");

        return markup;
    }

    function getAppointmentColor(resourceId) {
        return DevExpress.data.query(roomInfoList)
                .filter("id", resourceId)
                .toArray()[0].color;
    }

    function getAppointmentRoomName(resourceId) {
        return DevExpress.data.query(roomInfoList)
                .filter("id", resourceId)
                .toArray()[0].text;
    }
});

