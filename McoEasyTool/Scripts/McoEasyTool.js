function GetAccountsList(target, account_url) {
    $.ajax({
        url: SITENAME + account_url,
        type: "GET",
        success: function (data) {
            target.html(data);
        },
        error: function (error) {
            alert('Erreur lors de la récupération des comptes');
        },
        complete: function () { }
    });
}

function AccountEditor(id, target, url, origin, account_url) {
    login = $(".username").text();
    div = "<div class='mco-account-div'>";
    div = div + "<p>Veuillez choisir le compte d'utilisateur qui sera utilisé pour les actions sur cet objet</p>";
    div = div + "<p class='nav-info'>Vous devrez valider ces modifications avec votre mot de passe de session</p>";
    table = "<table class='mco-account-table'>"
    table = table + "<tr><td class='mco-account-label'>Compte d'exécution:</td><td class='mco-account-input'><select id='ExecutionAccount-" + id + "'></select></td></tr>";
    table = table + "<tr></tr>";
    table = table + "<tr><td class='mco-account-label'>Compte de check: </td><td class='mco-account-input'><select id='CheckAccount-" + id + "'></select></td></tr>";
    table = table + "<tr></tr>";
    table = table + "<tr><td class='mco-account-label'>Mot de passe " + login + "</td><td class='mco-account-input'><input type='password' id='password-" + id + "'/></td></tr></table>";
    div = div + table;
    div = div + "<p class='warning'>Si vous désirez plus de choix, vous devrez d'abord rajouter des comptes</p>";
    div = div + "</div>";
    var content = {
        state0: {
            title: "<span class='page-title-info warning'>Editeur des comptes</span>",
            html: div,
            buttons: { VALIDER: true },
            focus: 1,
            position: { container: "#container", x: 100, y: 0, width: 550 },
            submit: function (e, v, m, f) {
                execution_account = $("#ExecutionAccount-" + id).val();
                check_account = $("#CheckAccount-" + id).val();
                session_password = $("#password-" + id).val();
                if (session_password == null || session_password.trim() == "") {
                    alert("Le mot de passe ne peut être vide.\nEntrez le mot de passe de l'utilisateur " + login + " SVP.");
                }
                else {
                    target.prepend(loadinggif);
                    $.ajax({
                        url: SITENAME + url,
                        type: "POST",
                        data: { execution_account: execution_account, check_account: check_account, session_password: session_password },
                        success: function (data) {
                            alert(data);
                        },
                        error: function (error) {
                            alert('Erreur lors de la communication avec le serveur');
                        },
                        complete: function () {
                            Reload(target, SITENAME + origin)
                        }
                    });
                }
            }
        }
    };
    $.prompt(content);
    GetAccountsList($("#ExecutionAccount-" + id), account_url);
    GetAccountsList($("#CheckAccount-" + id), account_url);
}

function isNumber(input) {
    number = input.value;
    if (isNaN(number)) {
        input.value = 1;
    }
}

function get_server_disks_list(select, servername) {
    list = null;
    if (servername != null && servername != "") {
        $("body").toggleClass("wait");
        $.ajax({
            url: SITENAME + 'Servers/GetDisksList',
            type: "POST",
            data: { servername: servername },
            success: function (response) {
                list = response;
            },
            error: function (error) {
                alert(error.responseText);
            },
            complete: function () {
                if (list != null && list.trim() != "") {
                    disks = list.split("; ");
                    options = "";
                    for (index = 0; index < disks.length; index++) {
                        if (disks[index].trim() != "") {
                            options = options + "<option value='" + disks[index] + "'>Partition " +
                                disks[index] + "</option>";
                        }
                    }
                    if (select != null) {
                        select.html(options);
                    }
                }
                $("body").toggleClass("wait");
            }
        });
    }
}

function get_server_shares_list(select, servername) {
    list = null;
    if (servername != null && servername != "") {
        $("body").toggleClass("wait");
        $.ajax({
            url: SITENAME + 'Servers/GetSharesList',
            type: "POST",
            data: { servername: servername },
            success: function (response) {
                list = response;
            },
            error: function (error) {
                alert(error.responseText);
            },
            complete: function () {
                if (list != null && list.trim() != "") {
                    disks = list.split("; ");
                    options = "";
                    for (index = 0; index < disks.length; index++) {
                        if (disks[index].trim() != "") {
                            options = options + "<option value='" + disks[index] + "'>" +
                                disks[index] + "</option>";
                        }
                    }
                    if (select != null) {
                        select.html(options);
                    }
                }
                $("body").toggleClass("wait");
            }
        });
    }
}

function get_arrows(target, link, first_value, left_value, right_value, last_value) {

    enable_first = (first_value != 0) ? true : false;
    enable_left = (left_value != 0) ? true : false;
    enable_right = (right_value != 0) ? true : false;
    enable_last = (last_value != 0) ? true : false;

    first = "<img id='" + target.attr("id") + "_first-arrow' name='" + link + first_value + "' class='arrow first-arrow";
    left = "<img id='" + target.attr("id") + "_left-arrow' name='" + link + left_value + "' class='arrow left-arrow";
    right = "<img id='" + target.attr("id") + "_right-arrow' name='" + link + right_value + "' class='arrow right-arrow";
    last = "<img id='" + target.attr("id") + "_last-arrow' name='" + link + last_value + "' class='arrow last-arrow";

    if (enable_first) {
        first = first + " enabled-arrow' onclick='trigger_arrow_action(this);'";
    }
    else { first = first + " disabled-arrow'"; }
    if (enable_left) {
        left = left + " enabled-arrow' onclick='trigger_arrow_action(this);'";
    }
    else { left = left + " disabled-arrow'"; }
    if (enable_right) {
        right = right + " enabled-arrow' onclick='trigger_arrow_action(this);'";
    }
    else { right = right + " disabled-arrow'"; }
    if (enable_last) {
        last = last + " enabled-arrow' onclick='trigger_arrow_action(this);'";
    }
    else { last = last + " disabled-arrow'"; }

    first = first + "></img>";
    left = left + "></img>";
    right = right + "></img>";
    last = last + "></img>";

    arrows = "<table id='" + target.attr("id") + "-arrows' class='arrows'><tr>" +
        "<td>" + first + "</td>" +
        "<td></td><td></td>" +
        "<td>" + left + "</td>" +
        "<td></td><td></td>" +
        "<td>" + right + "</td>" +
        "<td></td><td></td>" +
        "<td>" + last + "</td>" +
    "</tr></table>";
    target.append(arrows);
}

function trigger_arrow_action(img) {
    target = img.id;
    target = target.substr(0, target.indexOf("_"));
    link = SITENAME + img.name;
    $.ajax({
        url: link,
        type: "POST",
        success: function (data) {
            $("#" + target).html(data);
        },
        error: function (error) {
            alert('Erreur lors de la communication avec le serveur');
        },
        complete: function () {
            colourTitles();
        }
    });
}

function get_reftech_status(stat) {
    stat = stat.trim();
    var status = "Inconnu";
    switch (stat.toUpperCase()) {
        case "O": status = "Opérationnel"; break;
        case "R": status = "Retiré"; break;
        case "A": status = "A venir"; break;
        default: break;
    }
    return status;
}

//PATH CHECKER
function isValidPath(path) {
    path = path.trim();
    if (path.indexOf(":") == 1) {
        parts = path.split("\\");
        if (parts.length > 1 && parts[1].trim() != "") {
            filename = parts[parts.length - 1].trim();
            extension = filename.split(".");
            if (filename.indexOf(".") != -1 && extension.length == 2 && extension[1].trim() != "") {
                return true;
            }
            else {
                return false;
            }
        }
        else {
            return false;
        }
    }
    else {
        return false;
    }
}
//END PATH CHECKER

//URL CHECKER
function isValidUrl(url) {
    url = url.trim();
    if (url.indexOf("http://") == 0 || url.indexOf("https://") == 0) {
        doubleurl = url.split("http");
        doubleurls = url.split("https");
        if (doubleurl.length == 2 || doubleurls.length == 2) {
            return true;
        }
        else {
            return false;
        }
    }
    else {
        return false;
    }
}

function correctUrl(url) {
    if (isValidUrl(url)) {
        return url;
    }
    else {
        return "http://" + url;
    }
}
//END URL CHECKER

if (typeof String.prototype.startsWith != 'function') {
    String.prototype.startsWith = function (str) {
        return this.substring(0, str.length) === str;
    }
};

if (typeof String.prototype.trim !== 'function') {
    String.prototype.trim = function () {
        return this.replace(/^\s+|\s+$/g, '');
    }
}

function AddressParser() {
    $("#input-recipient").keypress(function (key) {
        if (key.keyCode == 32 || key.keyCode == 59 || key.keyCode == 44) {
            $(this).val($(this).val() + "; ");
            return false;
        }
    });
}

function validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}

function colourTitles() {
    $(".page-title").each(function () {
        $(this).css("background-color", $(".active-nav-tab").css("background-color"));
    });;
}

function Reload(target, url) {
    $.ajax({
        url: url,
        type: "GET",
        success: function (data) {
            target.html(data);
        },
        error: function (error) {
            alert('Erreur lors de communication avec le serveur');
        },
        complete: function () { colourTitles(); }
    });
}

function select_tab(id) {
    var target_div = $("#" + id);
    $(".active-tab").each(function () {
        $(this).removeClass("active-tab");
        if (!$(this).hasClass("inactive-tab")) {
            $(this).addClass("inactive-tab");
        }
    });
    if (target_div.hasClass("inactive-tab")) {
        target_div.removeClass("inactive-tab");
    }
    if (!target_div.hasClass("active-tab")) {
        target_div.addClass("active-tab");
    }
    update_nav_tabs(id);
}

function launch_home(e, id) {
    tab = e.parentNode.getElementsByClassName("tab-title")[0];
    url = null;
    target = null;
    switch (id) {
        case "mco-ad-div": url = "McoAd/Home"; target = AD_MCO_DIV; break;
        case "mco-besr-div": url = "McoBesr/Home"; target = BESR_MCO_DIV; break;
        case "mco-app-div": url = "McoApp/Home"; target = APP_MCO_DIV; break;
        case "mco-space-div": url = "McoSpace/Home"; target = SPACE_MCO_DIV; break;
        default: url = "Home/Home"; target = COMMON_MCO_DIV; break;
    }
    $.ajax({
        url: SITENAME + url,
        type: "GET",
        success: function (data) {
            target.html(data);
        },
        error: function (error) {
            alert('Erreur lors de communication avec le serveur');
        },
        complete: function () {
            colourTitles();
            switch_tab(tab, id);
        }
    });
}

function switch_tab(e, id) {
    $(".active-nav-tab").each(function () {
        if ($(this).attr("id") != "tab-" + id) {
            select_tab(id);
            update_nav_tabs(id);
            colourTitles();
        }
    });
}

function close_tab(e, id) {
    length = $(".nav-tab").length;
    if (length > 1) {
        tab_id = "tab-" + id;
        targeted_tab = $("#" + tab_id);
        targeted_tab.remove();
        $("#" + id).empty();
        if ($(".active-nav-tab").length == 0) {
            successor = $(".nav-tab").eq(length - 2)
            successor.removeClass("inactive-nav-tab");
            successor.addClass("active-nav-tab");
            $(".active-tab").each(function () {
                $(this).removeClass("active-tab");
                if (!$(this).hasClass("inactive-tab")) {
                    $(this).addClass("inactive-tab");
                }
            });
            tab = successor.attr("id").substr(4);
            if ($("#" + tab).hasClass("inactive-tab")) {
                $("#" + tab).removeClass("inactive-tab");
            }
            if (!$("#" + tab).hasClass("active-tab")) {
                $("#" + tab).addClass("active-tab");
            }
            colourTitles();
        }
    }
    else {
        alert("Le dernier onglet n'est pas fermable.");
    }
    return false;
}

function update_nav_tabs(id) {
    var tab = $("#tab-" + id);
    $(".nav-tab").each(function () {
        if ($(this).hasClass("active-nav-tab")) {
            $(this).removeClass("active-nav-tab");
        }
        if (!$(this).hasClass("inactive-nav-tab")) {
            $(this).addClass("inactive-nav-tab");
        }
    });
    if (tab.length != 0) {
        tab.addClass("active-nav-tab");
    }
    else {
        var tab_icon = "<span class='tab-icon mco-", tab_leave = "<span class='tab-leave' onclick='close_tab(this,\"";
        var tab_name = "", tab_class = "";
        switch (id) {
            case "mco-ad-div": tab_name = "Active Di."; tab_class = "ad"; tab_icon += "ad"; break;
            case "mco-besr-div": tab_name = "Backup Ex."; tab_class = "besr"; tab_icon += "besr"; break;
            case "mco-app-div": tab_name = "Application"; tab_class = "app"; tab_icon += "app"; break;
            case "mco-space-div": tab_name = "C.Planning"; tab_class = "space"; tab_icon += "space"; break;
            default: tab_name = "General"; tab_class = "common"; tab_icon += "common"; break;
        }
        tab_icon += "'  onclick='launch_home(this,\"" + id + "\");' ></span>";
        tab_leave += id + "\");' title='fermer'></span>";
        var new_tab = "<div class='nav-tab " + tab_class + " active-nav-tab' id='tab-" + id + "'>" + tab_icon +
                            "<span  onclick='switch_tab(this,\"" + id + "\");' class='tab-title'>" + tab_name + "</span>" + tab_leave + "</div>";
        $("#nav-tabs").append(new_tab);
    }
}

function HasUnachieviedReport(controller, funct, args) {
    link = SITENAME + controller + "HasUnachieviedReport";
    erase = false;
    $("body").toggleClass("wait");
    $.ajax({
        url: link,
        type: "GET",
        success: function (data) {
            if (data != "OK") {
                erase = confirm(data);
            }
            else {
                funct(args);
            }
        },
        error: function (error) {
            alert('Erreur lors de la communication avec le serveur\n'
                + error.responseText);
        },
        complete: function () {
            if (erase) {
                CancelReport(controller, funct, args);
            }
            else {
                $("body").toggleClass("wait");
            }

        }
    });
}

function CancelReport(controller, funct, args) {
    deleted = "";
    $.ajax({
        url: SITENAME + controller + "Purge",
        type: "GET",
        success: function (data) {
            deleted = data;
        },
        error: function (error) {
            alert('Erreur lors de la communication avec le serveur');
        },
        complete: function () {
            $("body").toggleClass("wait");
            if (deleted.trim() == "") {
                deleted = "Aucun rapport supprimé";
            }
            alert(deleted);
            funct(args);
        }
    });
}


jQuery(document).ready(function () {


    //DropDown Menu
    var timeout = 500;
    var closetimer = 0;
    var ddmenuitem = 0;
    var subclosetimer = 0;
    var subddmenuitem = 0;

    function dropdownmenu_open() {
        dropdownmenu_canceltimer();
        dropdownmenu_close();
        ddmenuitem = $(this).find('#nav-ul-menu').css('visibility', 'visible');
    }

    function dropdownmenu_close()
    { if (ddmenuitem) ddmenuitem.css('visibility', 'hidden'); }

    function dropdownmenu_timer()
    { closetimer = window.setTimeout(dropdownmenu_close, timeout); }

    function dropdownmenu_canceltimer() {
        if (closetimer) {
            window.clearTimeout(closetimer);
            closetimer = null;
        }
    }

    $(document).ready(function () {
        $('#nav-ul > li').bind('mouseover', dropdownmenu_open)
        $('#nav-ul > li').bind('mouseout', dropdownmenu_timer)
    });

    document.onclick = dropdownmenu_close;

    //SUB MENU
    function dropdownsubmenu_open() {
        dropdownsubmenu_canceltimer();
        dropdownsubmenu_close();
        subddmenuitem = $(this).find('ul').css('visibility', 'visible');
    }

    function dropdownsubmenu_close()
    { if (subddmenuitem) subddmenuitem.css('visibility', 'hidden'); }

    function dropdownsubmenu_timer()
    { subclosetimer = window.setTimeout(dropdownsubmenu_close, timeout); }

    function dropdownsubmenu_canceltimer() {
        if (subclosetimer) {
            window.clearTimeout(subclosetimer);
            subclosetimer = null;
        }
    }

    $(document).ready(function () {
        $('#nav-ul-menu > li').bind('mouseover', dropdownsubmenu_open)
        $('#nav-ul-menu > li').bind('mouseout', dropdownsubmenu_timer)
    });

    document.onclick = dropdownsubmenu_close;

    function DateTranslator() {
        var MonthNames = ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin",
                        "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"];
        var DayNames = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi",
                        "Samedi"];
        var Today = new Date();
        Day = parseInt(Today.getDay());
        Month = parseInt(Today.getMonth());
        Year = Today.getFullYear();
        $("#header-account-date").text(DayNames[Day] + " " + Today.getDate() + " " + MonthNames[Month] + " " + Year);
    }

    function StartTime() {
        var Today = new Date();
        var Hours = Today.getHours();
        var Minutes = Today.getMinutes();
        var Seconds = Today.getSeconds();
        // add a zero in front of numbers<10
        Minutes = CheckTime(Minutes);
        Seconds = CheckTime(Seconds);
        $("#header-account-clock").html(Hours + ":" + Minutes + ":" + Seconds);
        DateTranslator();
        Timer = setTimeout(function () { StartTime() }, 500);
    }

    function CheckTime(Time) {
        if (Time < 10) {
            Time = "0" + Time;
        }
        return Time;
    }

    StartTime();

    $(".tab-leave").click(function (e) {
        alert(e.target);
    });

    $(".inactive-nav-btn").click(function InactiveStylingButton() {
        PreviousButton = $(".active-nav-btn");
        PreviousButton.removeClass("active-nav-btn");
        PreviousButton.addClass("inactive-nav-btn");
        $(this).removeClass("inactive-nav-btn");
        $(this).addClass("active-nav-btn");
    });

    $(".li-inactive-nav-btn").click(function LiInactiveStylingButton() {
        if ($(this).hasClass("act")) {
            PreviousButton = $(".active-nav-btn");
            PreviousButton.removeClass("active-nav-btn");
            PreviousButton.addClass("inactive-nav-btn");
        }

        if ($(this).hasClass("sub")) {
            PreviousLiButton = $(".li-active-nav-btn");
            PreviousLiButton.removeClass("li-active-nav-btn");
            PreviousLiButton.addClass("li-inactive-nav-btn");
            parentUl = $(this).closest("ul").parent().find(".act");
            parentUl.removeClass("li-inactive-nav-btn");
            parentUl.addClass("li-active-nav-btn");
        }

        $("#nav-check-btn").find(".inactive-nav-btn").trigger("click");
        $(this).addClass("li-active-nav-btn");
    });

    $("body.wait").click(function (e) {
        e.preventDefault();
        return false;
    });

    $("#scan-error-detector").click(function ScanErrorDetector() {
        password = "";
        domain = "";
        username = "";
        login = $(".username").text();
        login = login.replace("\\", " ");
        message = "<div>Cette fonction est utilisée pour la correction de l'erreur 8606 pour les serveurs du dernier " +
                    "rapport qui a été effectué. Plusieurs copies seront effectuées à travers le réseau sur des serveurs distants. " +
                    "Il se peut que celles-ci ne s'effectuent pas correctement si votre compte ne dispose pas des droits d'accès requis " +
                    "pour ces opérations.</div><br />" +
                    "ENTREZ LES INFORMATIONS D'UN COMPTE UTILISATEUR VALIDE SVP <br />";
        domaindiv = "<label>Domaine: </label><input type='text' value='" + login.split(" ")[0] + "' class='domain-input' id='domain'/><br />";
        usernamediv = "<label>Nom d'utilisateur: </label><input type='text' value='" + login.split(" ")[1] + "'class='username-input' id='username'/><br />";

        var contenu = {
            state0: {
                title: "Correction Erreur 8606",
                html: message + domaindiv + usernamediv + "<label>Mot de passe: </label><input type='password' class='password-input' id='password'/>",
                buttons: { VALIDER: true },
                focus: 1,
                position: { container: "#container", x: 100, y: 0, width: 450 },
                submit: function (e, v, m, f) {
                    domain = $("#domain").val();
                    username = $("#username").val();
                    password = $("#password").val();
                    AD_MCO_DIV.prepend(loadinggif);
                    $.ajax({
                        url: SITENAME + 'McoAd/LaunchTestError',
                        type: "POST",
                        data: { domain: domain, username: username, password: password },
                        success: function (data) {
                            alert(data);
                        },
                        error: function (error) {
                            alert('Erreur lors de la communication avec le serveur');
                        },
                        complete: function () {
                            NAV_AD_REPORTS_BTN.trigger("click");
                        }
                    });
                }
            }
        };
        $.prompt(contenu);
    });

    $(".mco-common-functions").click(function () {
        url = SITENAME + "Home/";
        message = "Erreur de communication serveur";
        var identity = $(this).attr("id");
        COMMON_MCO_DIV.prepend(loadinggif);
        switch (identity) {
            case "nav-home-btn": url = url + "Home"; break;
            case "nav-contact-btn": url = url + "Contact"; break;
            case "nav-tools-btn": url = url + "Utilities"; break;
            case "nav-accounts-btn": url = SITENAME + "Accounts/DisplayAccounts"; break;
            default: break;
        }
        select_tab("mco-common-div");
        $.ajax({
            url: url,
            type: "GET",
            success: function (data) {
                COMMON_MCO_DIV.html(data);
            },
            error: function (error) {
                COMMON_MCO_DIV.html(error.responseText);
                //NAV_HOME_BTN.trigger("click");
            },
            complete: function () { colourTitles(); }
        });
    });

    $(".mco-ad-functions").click(function () {
        url = SITENAME + "McoAd/";
        message = "Erreur de communication serveur";
        var identity = $(this).attr("id");
        AD_MCO_DIV.prepend(loadinggif);
        switch (identity) {
            case "scan-scheduler": url = url + "DisplaySchedules"; break;
            case "nav-reports-btn": url = url + "DisplayReports"; break;
            case "nav-faultyservers-btn": url = url + "DisplayFaultyServers"; break;
            case "nav-adsettings-btn": url = url + "DisplaySettings"; break;
            case "nav-recipients-btn": url = url + "DisplayRecipients"; break;
            default: url = url + "Home"; break;
        }
        select_tab("mco-ad-div");
        $.ajax({
            url: url,
            type: "GET",
            success: function (data) {
                AD_MCO_DIV.html(data);
            },
            error: function (error) {
                AD_MCO_DIV.html(error.responseText);
                //NAV_HOME_BTN.trigger("click");
            },
            complete: function () { colourTitles(); }
        });
    });

    $(".mco-besr-functions").click(function () {
        url = SITENAME + "McoBesr/";
        message = "Erreur de communication serveur";
        var identity = $(this).attr("id");
        BESR_MCO_DIV.prepend(loadinggif);
        switch (identity) {
            case "nav-besrreportsinit-btn": url = url + "DisplayChecker"; break;
            case "nav-besrreports-btn": url = url + "DisplayReports"; break;
            case "besr-scheduler": url = url + "DisplaySchedules"; break;
            case "nav-pools-btn": url = url + "DisplayPools"; break;
            case "failed-backupservers-btn": url = url + "DisplayFailedServers"; break;
            case "init-pools-btn": url = url + "DisplayImporter"; break;
            case "nav-besrrecipients-btn": url = url + "DisplayRecipients"; break;
            default: url = url + "Home"; break;
        }
        select_tab("mco-besr-div");
        $.ajax({
            url: url,
            type: "GET",
            success: function (data) {
                BESR_MCO_DIV.html(data);
            },
            error: function (error) {
                BESR_MCO_DIV.html(error.responseText);
                //NAV_HOME_BTN.trigger("click");
            },
            complete: function () { colourTitles(); }
        });
    });

    $(".mco-app-functions").click(function () {
        url = SITENAME + "McoApp/";
        message = "Erreur de communication serveur";
        var identity = $(this).attr("id");
        APP_MCO_DIV.prepend(loadinggif);
        switch (identity) {
            case "nav-appreportsinit-btn": url = url + "DisplayChecker"; break;
            case "nav-applications-btn": url = url + "DisplayApplications"; break;
            case "app-scheduler": url = url + "DisplaySchedules"; break;
            case "nav-appreports-btn": url = url + "DisplayReports"; break;
            case "failed-applications-btn": url = url + "DisplayFailedApplications"; break;
            case "init-applications-btn": url = url + "DisplayImporter"; break;
            case "nav-apprecipients-btn": url = url + "DisplayRecipients"; break;
            case "nav-appdomains-btn": url = url + "DisplayAppDomains"; break;
            case "nav-detailledappreportsinit-btn": url = url + "DisplayFurhterChecker"; break;
            default: url = url + "Home"; break;
        }
        select_tab("mco-app-div");
        $.ajax({
            url: url,
            type: "GET",
            success: function (data) {
                APP_MCO_DIV.html(data);
            },
            error: function (error) {
                APP_MCO_DIV.html(error.responseText);
                //NAV_HOME_BTN.trigger("click");
            },
            complete: function () { colourTitles(); }
        });
    });

    $(".mco-space-functions").click(function () {
        url = SITENAME + "McoSpace/";
        message = "Erreur de communication serveur";
        var identity = $(this).attr("id");
        SPACE_MCO_DIV.prepend(loadinggif);
        switch (identity) {
            //case "nav-spacereportsinit-btn": url = url + "DisplayChecker"; break;
            case "nav-spacereports-btn": url = url + "DisplayReports"; break;
            case "space-scheduler": url = url + "DisplaySchedules"; break;
            case "nav-spaces-btn": url = url + "DisplaySpaceServers"; break;
            case "init-spaces-btn": url = url + "DisplayImporter"; break;
            case "nav-spacerecipients-btn": url = url + "DisplayRecipients"; break;
            default: url = url + "Home"; break;
        }
        select_tab("mco-space-div");
        $.ajax({
            url: url,
            type: "GET",
            success: function (data) {
                SPACE_MCO_DIV.html(data);
            },
            error: function (error) {
                SPACE_MCO_DIV.html(error.responseText);
                //NAV_HOME_BTN.trigger("click");
            },
            complete: function () { colourTitles(); }
        });
    });

    function LeaveConfirm() {
        window.onbeforeunload = function () {
            return "Si vous quittez cette page maintenant, " +
                            "l'email ne sera pas envoyé!!!!";
        };
    }

    function AddressParser() {
        $("#input-recipient").keypress(function (key) {
            if (key.keyCode == 32 || key.keyCode == 59 || key.keyCode == 44) {
                $(this).val($(this).val() + "; ");
                return false;
            }
        });
    }

    $("#nav-spacereportsinit-btn").click(function Check() {
        scannow = confirm("Voulez-vous lancer le Check Capacity Planning maintenant?\n L'analyse se déroulera pendant un certain temps");
        if (!scannow) {
            return false;
        }
        var args = {}
        HasUnachieviedReport("McoSpace/", SpaceLaunch, args);
    });

    function SpaceLaunch(args) {
        SPACE_MCO_DIV.prepend(loadinggif);
        $.ajax({
            url: SITENAME + 'McoSpace/CheckCapacityPlanning',
            type: "GET",
            success: function (data) {
                emailId = data.email;
                errors = data.errors;
                alert("Le processus s'est déroulé avec les erreurs suivantes:\n" + errors);
                email = "<div id='ReadyEmail' class='ReadyEmail fade'>" +
                                "<label class='ReadyEmail-label'>Destinataires: </label>" +
                                " <input type='text' id='input-recipient'class='ReadyEmail-input'/><br />" +
                                "<input type='button' id='input-button' onClick='SendSpaceEmail(" + emailId + ");'class='ReadyEmail-submit' value='Envoyer' />" +
                                "<label class='ReadyEmail-label'>Sujet : </label>" +
                                "<input type='text' id='input-subject'  class='ReadyEmail-input'/><br />" +
                                "<div id='input-body'  class='ReadyEmail-body'>" +
                                "<textarea id='input-message' placeholder='Ajoutez un message ici...' " +
                                "style='margin-left:10px;width:800px;border:0px solid #000;'" +
                                "class='ReadyEmail-input'></textarea></div></div>";
                SPACE_MCO_DIV.html(email);
                AddressParser();
                $.ajax({
                    url: SITENAME + 'Emails/Open/' + emailId,
                    dataType: "JSON",
                    type: "GET",
                    success: function (content) {
                        SPACE_MCO_DIV.find("#input-recipient").eq(0).val(content.recipient);
                        SPACE_MCO_DIV.find("#input-subject").eq(0).val(content.subject);
                        SPACE_MCO_DIV.find("#input-body").eq(0).append(content.body);
                    },
                    error: function (error) {
                        alert("Erreur lors de l'ouverture de l'email");
                    },
                    complete: function () { }
                });
            },
            error: function (error) {
                alert(error.responsText);
                NAV_HOME_BTN.trigger("click");
            }
        });
    }

    $("#scan-launcher").click(function Check() {
        scannow = confirm("Voulez-vous lancer le Check AD maintenant?\n L'analyse prendra une vingtaine de minutes");
        if (!scannow) {
            return false;
        }
        var args = {}
        HasUnachieviedReport("McoAd/", AdLaunch, args);
    });

    function AdLaunch(args) {
        AD_MCO_DIV.prepend(loadinggif);
        $.ajax({
            url: SITENAME + 'McoAd/CheckActiveDirectory',
            type: "GET",
            success: function (data) {
                reportId = data.Report;
                emailId = data.Email;
                errors = data.Errors;
                alert("Le processus s'est déroulé avec les erreurs suivantes:\n" + errors);
                email = "<div id='ReadyEmail' class='ReadyEmail fade'>" +
                                "<label class='ReadyEmail-label'>Destinataires: </label>" +
                                " <input type='text' id='input-recipient'class='ReadyEmail-input'/><br />" +
                                "<input type='button' id='input-button' onClick='SendEmail(" + emailId + ");'class='ReadyEmail-submit' value='Envoyer' />" +
                                "<label class='ReadyEmail-label'>Sujet : </label>" +
                                "<input type='text' id='input-subject'  class='ReadyEmail-input'/><br />" +
                                "<div id='input-body'  class='ReadyEmail-body'>" +
                                "<textarea id='input-message' placeholder='Ajoutez un message ici...' " +
                                "style='margin-left:10px;width:800px;border:0px solid #000;'" +
                                "class='ReadyEmail-input'></textarea></div></div>";
                AD_MCO_DIV.html(email);
                AddressParser();
                $.ajax({
                    url: SITENAME + 'Emails/Open/' + emailId,
                    dataType: "JSON",
                    type: "GET",
                    success: function (content) {
                        AD_MCO_DIV.find("#input-recipient").eq(0).val(content.recipient);
                        AD_MCO_DIV.find("#input-subject").eq(0).val(content.subject);
                        AD_MCO_DIV.find("#input-body").eq(0).append(content.body);
                    },
                    error: function (error) {
                        alert("Erreur lors de l'ouverture de l'email");
                    },
                    complete: function () { }
                });
            },
            error: function (error) {
                alert(error.responsText);
                NAV_HOME_BTN.trigger("click");
            }
        });
    }
});


function SendEmail(EmailId) {
    select_tab("mco-ad-div");
    Recipient = $("#input-recipient").val();
    Subject = $("#input-subject").val();
    Message = $("#input-message").val();
    $.ajax({
        url: SITENAME + 'Emails/Send/' + EmailId,
        dataType: "JSON",
        type: "POST",
        data: { Subject: Subject, Recipient: Recipient, Message: Message },
        success: function (response) {
            alert(response.Email + "\n" + response.Response);
        },
        error: function (error) {
            alert("Erreur lors de l'envoi du mail");
        },
        complete: function () {
            window.onbeforeunload = null;
            autocorrect = confirm("Voulez-vous lancer la correction de l'erreur 8606 pour ce dernier rapport maintenant?");
            if (autocorrect) {
                $("#scan-error-detector").trigger("click");
            }
            else {
                NAV_AD_REPORTS_BTN.trigger("click");
            }
        }
    });
}

function SendBackupEmail(EmailId, autobesr) {
    select_tab("mco-besr-div");
    Recipient = $("#input-recipient").val();
    Subject = $("#input-subject").val();
    Message = $("#input-message").val();
    $.ajax({
        url: SITENAME + 'Emails/Send/' + EmailId,
        dataType: "JSON",
        type: "POST",
        data: { Subject: Subject, Recipient: Recipient, Message: Message },
        success: function (response) {
            alert(response.BackupEmail + "\n" + response.Response);
        },
        error: function (error) {
            alert("Erreur lors de l'envoi du mail");
        },
        complete: function () {
            window.onbeforeunload = null;
            if (!autobesr) {
                manage = confirm("Voulez-vous gérer les serveurs en erreur de sauvegarde?");
                if (manage) {
                    $("#autobesr-div").css("visibility", "visible");
                    $("#autobesr-div").css("display", "block");
                    BESR_MCO_DIV.html($("#autobesr-div").html());
                    //----------------------------------------------------------
                    $(".deleter").click(function () {
                        deletor = confirm("Voulez vous réellement supprimer ce serveur de la liste?");
                        line = $(this).closest("tr");
                        serverId = $(this).closest("td").find(".action-id-getter").val();
                        if (deletor) {
                            BESR_MCO_DIV.prepend(loadinggif);
                            $.ajax({
                                url: SITENAME + 'McoBesr/DeleteBackupServer/' + serverId,
                                type: "GET",
                                success: function (response) {
                                    alert(response);
                                },
                                error: function (error) {
                                    alert(error.responseText);
                                },
                                complete: function () {
                                    line.remove();
                                    $(".loading-gif-div").remove();
                                }
                            });
                        }
                    });

                    $(".backup-launcher").click(function () {
                        launch = confirm("Voulez vous vraiment lancer la sauvegarde pour ce serveur maintenant? \n" +
                                       "Les anciens fichiers de sauvegardes de ce serveur seront éventuellement supprimés.");
                        if (launch) {
                            serverId = $(this).closest("td").find(".action-id-getter").val();
                            line = $(this).closest("tr");
                            success = false;
                            BESR_MCO_DIV.prepend(loadinggif);
                            $.ajax({
                                url: SITENAME + 'McoBesr/ServiceLauncher/' + serverId,
                                type: "GET",
                                success: function (content) {
                                    if ((content.indexOf('Backup Exec System Recovery: Running: Manual') != -1) &&
                                                   (content.indexOf('SymSnapService: Running: Manual') != -1)
                                               ) {
                                        line.find("td").eq(5).text("OK");
                                        success = true;
                                    }
                                },
                                error: function (error) {
                                    alert(error.responseText);
                                },
                                complete: function () {
                                    if (success) {
                                        $.ajax({
                                            url: SITENAME + 'McoBesr/BackupExecLauncher/' + serverId,
                                            type: "GET",
                                            success: function (content) {
                                                line.find("td").eq(6).text("Relancé");
                                                alert(content);
                                            },
                                            error: function (error) {
                                                alert(error.responseText);
                                            },
                                            complete: function () {
                                                $(".loading-gif-div").remove();
                                                line.find("td").eq(0).empty();
                                                line.find("td").eq(8).empty();
                                            }
                                        });
                                    }
                                }
                            });
                        }
                        else {
                            return false;
                        }
                    });

                    $(".backup-launcher-btn").click(function () {
                        selectedServers = "";

                        $(".selected-servers").each(function () {
                            if ($(this).is(':checked')) {
                                selectedServers += $(this).closest("td").find(".action-id-getter").val() + ", ";
                            }
                        });

                        if (selectedServers == "") {
                            alert("Vous devez sélectionner au moins un serveur");
                            return false;
                        }
                        else {
                            BESR_MCO_DIV.prepend(loadinggif);
                            $.ajax({
                                url: SITENAME + 'McoBesr/BackupExecLauncherServers/',
                                type: "POST",
                                data: { selectedServers: selectedServers },
                                success: function (content) {
                                    alert(content);
                                },
                                error: function (error) {
                                    alert(error.responseText);
                                },
                                complete: function () {
                                    NAV_BESR_CHECKS_INIT_BTN.trigger("click");
                                }
                            });
                        }
                    });
                }
                else {
                    NAV_BESR_REPORTS_BTN.trigger("click");
                }
            }
            else {
                NAV_BESR_REPORTS_BTN.trigger("click");
            }
        }
    });
}

function SendSpaceEmail(EmailId) {
    select_tab("mco-space-div");
    Recipient = $("#input-recipient").val();
    Subject = $("#input-subject").val();
    Message = $("#input-message").val();
    $.ajax({
        url: SITENAME + 'Emails/Send/' + EmailId,
        dataType: "JSON",
        type: "POST",
        data: { Subject: Subject, Recipient: Recipient, Message: Message },
        success: function (response) {
            alert(response.Email + "\n" + response.Response);
        },
        error: function (error) {
            alert("Erreur lors de l'envoi du mail");
        },
        complete: function () {
            window.onbeforeunload = null;
            NAV_SPACE_REPORTS_BTN.trigger("click");
        }
    });
}

function SendAppEmail(EmailId) {
    select_tab("mco-app-div");
    Recipient = $("#input-recipient").val();
    Subject = $("#input-subject").val();
    Message = $("#input-message").val();
    $.ajax({
        url: SITENAME + 'Emails/Send/' + EmailId,
        dataType: "JSON",
        type: "POST",
        data: { Subject: Subject, Recipient: Recipient, Message: Message },
        success: function (response) {
            alert(response.Email + "\n" + response.Response);
        },
        error: function (error) {
            alert("Erreur lors de l'envoi du mail");
        },
        complete: function () {
            window.onbeforeunload = null;
            NAV_APP_REPORTS_BTN.trigger("click");
        }
    });
}
