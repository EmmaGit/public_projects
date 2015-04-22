scheduleId = 0;
currentDisplayedMonth = 0;
actualMonth = 0;

function TimeFiller(hourselector, minuteselector) {
    var hours = "";
    var minutes = "";
    for (hour = 0; hour < 24; hour++) {
        if (hour < 10) {
            hours += "<option id='hour-0" + hour + "'>0" + hour + "</option>";
        }
        else {
            hours += "<option id='hour-" + hour + "'>" + hour + "</option>";
        }
    }
    hourselector.html(hours);

    for (minute = 0; minute < 60; minute++) {
        if (minute < 10) {
            minutes += "<option id='minute-0" + minute + "'>0" + minute + "</option>";
        }
        else {
            minutes += "<option id='hour-" + minute + "'>" + minute + "</option>";
        }
    }
    minuteselector.html(minutes);
}

function CalendarBuilder(thismonthdiv) {
    MonthNames = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin",
                        "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"];
    columnindex = 0;
    today = new Date();
    day = today.getDate();
    month = today.getMonth();
    year = today.getFullYear();
    lastday = new Date(year, month + 1, 0);
    thismonth = "<span id='calendar-table-title'>" + MonthNames[month] + "</span>";
    thismonth += "<table id='calendar-table' class='calendar-table'><tbody>";
    for (index = 1; index <= lastday.getDate() ; index++) {
        if (columnindex == 0) {
            thismonth += "<tr class='calendar-table-tr'>";
        }
        if (index < day) {
            thismonth += "<td class='calendar-table-td passedDay' name='" + month + "'>" + index + "</td>";
        }
        else {
            if (index == day) {
                thismonth += "<td class='calendar-table-td today thismonth' name='" + month + "'>" + index + "</td>";
            }
            else {
                thismonth += "<td class='calendar-table-td thismonth' name='" + month + "'>" + index + "</td>";
            }
        }
        columnindex++;
        if (columnindex == 6) {
            thismonth += "</tr>";
            columnindex = 0;
        }
    }
    actualMonth = month;
    currentDisplayedMonth = month;
    thismonth += "</tbody></table>";
    thismonthdiv.html(thismonth);
}

function NextMonthBuilder() {
    MonthNames = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin",
                        "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"];
    columnindex = 0;
    today = new Date();
    day = today.getDate();
    month = today.getMonth();
    year = today.getFullYear();
    MonthCalendar = "";
    for (index = month + 1; index <= 11; index++) {
        lastday = new Date(year, index + 1, 0);
        MonthCalendar += "<div class='nextmonth-div hidden-month' id='nextmonth-div-" + index + "'><span>" + MonthNames[index] + "</span>";
        MonthCalendar += "<table id='next-calendar-table-" + index + "' class='calendar-table'>";
        MonthCalendar += "<tbody>";
        for (dayIndex = 1; dayIndex <= lastday.getDate() ; dayIndex++) {
            if (columnindex == 0) {
                MonthCalendar += "<tr class='calendar-table-tr'>";
            }
            MonthCalendar += "<td class='calendar-table-td nextmonth' name='" + index + "'>" + dayIndex + "</td>";
            columnindex++;
            if (columnindex == 6) {
                MonthCalendar += "</tr>";
                columnindex = 0;
            }
        }
        MonthCalendar += "</tbody></table></div>";
        columnindex = 0;
    }
    $(".displayedmonth-div").append(MonthCalendar);
}

function TriggerTimeChanges(module) {
    //-----------------------------------------------
    $("#" + module + "_hourselector").change(function () {

        $("#" + module + "_hourselectorinput").val($(this).val());
    });

    //-----------------------------------------------
    $("#" + module + "_minuteselector").change(function () {
        $("#" + module + "_minuteselectorinput").val($(this).val());
    });

    //-----------------------------------------------
    $("#" + module + "_hourselectorinput, #" + module + "minuteselectorinput").keypress(function (key) {
        if (!(key.keyCode >= 48 && key.keyCode <= 57)) {
            return false;
        }
    });

    //-----------------------------------------------
    $("#" + module + "_hourselectorinput").blur(function () {
        if ($(this).val() > 23) {
            $(this).val("00");
        }
        $("#" + module + "_hourselector").prop("selectedIndex", $(this).val());
    });

    //-----------------------------------------------
    $("#" + module + "_minuteselectorinput").blur(function () {
        if ($(this).val() > 59) {
            $(this).val("00");
        }
        $("#" + module + "_minuteselector").prop("selectedIndex", $(this).val());
    });
}

$("#next-month").click(function () {
    if (currentDisplayedMonth < 11) {
        $(".displayedmonth-div div").each(function () {
            if (!($(this).hasClass("hidden-month"))) {
                $(this).addClass("hidden-month");
            }
            if (($(this).hasClass("displayed-month"))) {
                $(this).removeClass("displayed-month");
            }
        });
        currentDisplayedMonth++;
        $("#nextmonth-div-" + currentDisplayedMonth).removeClass("hidden-month");
        $("#nextmonth-div-" + currentDisplayedMonth).addClass("displayed-month");
    }
});

$("#previous-month").click(function () {
    if (currentDisplayedMonth > actualMonth + 1) {
        $(".displayedmonth-div div").each(function () {
            if (!($(this).hasClass("hidden-month"))) {
                $(this).addClass("hidden-month");
            }
            if (($(this).hasClass("displayed-month"))) {
                $(this).removeClass("displayed-month");
            }
        });
        currentDisplayedMonth--;
        $("#nextmonth-div-" + currentDisplayedMonth).removeClass("hidden-month");
        $("#nextmonth-div-" + currentDisplayedMonth).addClass("displayed-month");
    }
    else {
        if (currentDisplayedMonth == actualMonth + 1) {
            $(".displayedmonth-div div").each(function () {
                if (!($(this).hasClass("hidden-month"))) {
                    $(this).addClass("hidden-month");
                }
                if (($(this).hasClass("displayed-month"))) {
                    $(this).removeClass("displayed-month");
                }
            });
            currentDisplayedMonth--;
            $(".thismonth-div").removeClass("hidden-month");
            $(".thismonth-div").addClass("displayed-month");
        }
    }
});