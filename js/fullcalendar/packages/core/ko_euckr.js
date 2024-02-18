(function (global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) :
    (global = global || self, (global.FullCalendarLocales = global.FullCalendarLocales || {}, global.FullCalendarLocales.ko = factory()));
}(this, function () { 'use strict';

    var ko = {
        code: "ko",
        buttonText: {
            prev: "������",
            next: "������",
            today: "����",
            month: "��",
            week: "��",
            day: "��",
            list: "�������"
        },
        weekLabel: "��",
        allDayText: "����",
        eventLimitText: "��",
        noEventsMessage: "������ �����ϴ�"
    };

    return ko;

}));
