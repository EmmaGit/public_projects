﻿<%@ Page Language="C#" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage
<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    Home
</asp:Content>

<asp:Content ID="pl" ContentPlaceHolderID="MainContent" runat="server">
    <div id="header">
            
            <div class="header-logo-div">      
                <%@ Html.ActionLink("Mco", "Index", null, new { @class = "header-logo-img" })%>
            </div>

            <div id="nav-menu">
                <ul id="nav-ul" class="nav-ul">
                    <li id="nav-home-btn" class="nav-ul-li mco-common-functions">
                        <span class="nav-ul-li-span inactive-nav-btn">
                            <strong>
                                <a href="javascript:" class="nav-ul-li-a">Accueil</a>
                            </strong>
                        </span>
                    </li>

                    <li id="nav-check-btn" class="nav-ul-li">
                        <span class="nav-ul-li-span inactive-nav-btn">
                            <strong>
                                <a href="javascript:" class="nav-ul-li-a">Actes MCO</a>
                            </strong>
                        </span>
                        <ul id="nav-ul-menu" class="nav-ul-a">
                            <li id="mco-check-ad" class="nav-ul-li"
                                title="Check de réplication Active Directory.">
                                <span class="li-nav-ul-li-span li-inactive-nav-btn act">
                                    <strong>
                                        <a href="javascript:" class="nav-ul-li-a">Active Dir.</a>
                                    </strong>
                                </span>
                                <ul id="check-ad-ul" class="nav-sub-ul-a">
                                    <li id="scan-launcher" class="nav-ul-li mco-ad-functions"
                                        title="Vérifier l'état des mises à jour sur l'Active Directory Maintenant!!">
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Check AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="scan-error-detector" class="nav-ul-li mco-ad-functions"
                                        title="Lancer la détection de l'erreur 8606 pour le dernier rapport généré.">
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Détecter Erreur</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="scan-scheduler" class="nav-ul-li mco-ad-functions"
                                        title="Planifier le lancement du processus de vérification pour une date ultérieure.">
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Planification AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-reports-btn" class="nav-ul-li mco-ad-functions"
                                        title="Consultez l'historique des opérations." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Historique AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-faultyservers-btn" class="nav-ul-li mco-ad-functions"
                                        title="Consultez la liste des serveurs défaillants." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Controleurs KO AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-adsettings-btn" class="nav-ul-li mco-ad-functions"
                                        title="Gérer les paramètres de filtrage pour les serveurs en état critique." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Réglages AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-recipients-btn" class="nav-ul-li mco-ad-functions"
                                        title="Gérer les adresses qui seront chargées automatiquement pour ce module." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Destinataires AD</a>
                                            </strong>
                                        </span>
                                    </li>
                                </ul>
                            </li>

                            <li id="mco-check-besr" class="nav-ul-li"
                                title="Check de sauvegardes Windows.">
                                <span class="li-nav-ul-li-span li-inactive-nav-btn act">
                                    <strong>
                                        <a href="javascript:" class="nav-ul-li-a">Backup Exec.</a>
                                    </strong>
                                </span>
                                <ul id="check-besr-ul" class="nav-sub-ul-a">
                                    <li id="nav-besrreportsinit-btn" class="nav-ul-li mco-besr-functions"
                                        title="Vérifier l'état des sauvegardes au niveau des serveurs du jour et relancer les sauvegardes en echec." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Check BESR</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="besr-scheduler" class="nav-ul-li mco-besr-functions"
                                        title="Planifier d'avance une execution du check BESR." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Planification BESR</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-besrreports-btn" class="nav-ul-li mco-besr-functions"
                                        title="Consultez l'historique des opérations BESR." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Historique BESR</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="failed-backupservers-btn" class="nav-ul-li mco-besr-functions"
                                        title="Consultez la liste des serveurs en défaillance de sauvegarde pour la semaine en cours." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Serveurs KO</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-pools-btn" class="nav-ul-li mco-besr-functions"
                                        title="Gérer les différents pools et leurs serveurs." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Pools</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="init-pools-btn" class="nav-ul-li mco-besr-functions"
                                        title="Réinitialiser la liste des pools et leurs serveurs." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Réinitialiser BESR</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-besrrecipients-btn" class="nav-ul-li mco-besr-functions"
                                        title="Gérer les adresses qui seront chargées automatiquement pour ce module." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Destinataires BESR</a>
                                            </strong>
                                        </span>
                                    </li>
                                </ul>
                            </li>

                            <li id="mco-check-app" class="nav-ul-li"
                                title="Check des applications.">
                                <span class="li-nav-ul-li-span li-inactive-nav-btn act">
                                    <strong>
                                        <a href="javascript:" class="nav-ul-li-a">Flash Tests</a>
                                    </strong>
                                </span>
                                <ul id="check-app-ul" class="nav-sub-ul-a">
                                    <li id="nav-appreportsinit-btn" class="nav-ul-li mco-app-functions"
                                        title="Vérifier l'état des applications." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Check APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-detailledappreportsinit-btn" class="nav-ul-li mco-app-functions"
                                        title="Vérifier l'état des applications en mode détaillé (Connexion URL des serveurs)." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Vérification Détaillée</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="app-scheduler" class="nav-ul-li mco-app-functions"
                                        title="Planifier d'avance une execution du check APP." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Planification APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-appreports-btn" class="nav-ul-li mco-app-functions"
                                        title="Consultez l'historique des opérations APP." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Historique APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="failed-applications-btn" class="nav-ul-li mco-app-functions"
                                        title="Consultez la liste des applications down." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Applications KO</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-applications-btn" class="nav-ul-li mco-app-functions"
                                        title="Consulter la liste des applications." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Applications</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="init-applications-btn" class="nav-ul-li mco-app-functions"
                                        title="Réinitialiser la liste des applications." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Réinitialiser APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-apprecipients-btn" class="nav-ul-li mco-app-functions"
                                        title="Gérer les adresses qui seront chargées automatiquement pour ce module." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Destinataires APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-appdomains-btn" class="nav-ul-li mco-app-functions"
                                        title="Gérer les domaines d'applications." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Domaines APP</a>
                                            </strong>
                                        </span>
                                    </li>
                                </ul>
                            </li>

                            <li id="mco-check-space" class="nav-ul-li"
                                title="Check d'espaces disques.">
                                <span class="li-nav-ul-li-span li-inactive-nav-btn act">
                                    <strong>
                                        <a href="javascript:" class="nav-ul-li-a">Capacity Pln.</a>
                                    </strong>
                                </span>
                                <ul id="check-space-ul" class="nav-sub-ul-a">
                                    <li id="nav-spacereportsinit-btn" class="nav-ul-li mco-space-functions"
                                        title="Vérifier l'espace manquant sur les serveurs." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Check C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="space-scheduler" class="nav-ul-li mco-space-functions"
                                        title="Planifier d'avance une execution du check SPACE." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Planification C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-spacereports-btn" class="nav-ul-li mco-space-functions"
                                        title="Consultez l'historique des opérations SPACE." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Historique C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-spaces-btn" class="nav-ul-li mco-space-functions"
                                        title="Gérer les serveurs." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Serveurs C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="init-spaces-btn" class="nav-ul-li mco-space-functions"
                                        title="Réinitialiser la liste des serveurs." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Réinitialiser C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                    <li id="nav-spacerecipients-btn" class="nav-ul-li mco-space-functions"
                                        title="Gérer les adresses qui seront chargées automatiquement pour ce module." >
                                        <span class="li-nav-ul-li-span li-inactive-nav-btn sub">
                                            <strong>
                                                <a href="javascript:" class="nav-ul-li-a">Destinataires C.PLN</a>
                                            </strong>
                                        </span>
                                    </li>
                                </ul>
                            </li>

                        </ul>
                    </li>
                    <li id="nav-accounts-btn" class="nav-ul-li mco-common-functions">
                        <span class="nav-ul-li-span inactive-nav-btn">
                            <strong>
                                <a href="javascript:" class="nav-ul-li-a">Comptes</a>
                            </strong>
                        </span>
                    </li>
                    <li id="nav-tools-btn" class="nav-ul-li mco-common-functions">
                        <span class="nav-ul-li-span inactive-nav-btn">
                            <strong>
                                <a href="javascript:" class="nav-ul-li-a">Outils</a>
                            </strong>
                        </span>
                    </li>
                    <li id="nav-contact-btn" class="nav-ul-li mco-common-functions">
                        <span class="nav-ul-li-span inactive-nav-btn">
                            <strong>
                                <a href="javascript:" class="nav-ul-li-a">Contacts</a>
                            </strong>
                        </span>
                    </li>
                </ul>
            </div>
            
            <div class="header-account-div">
                <section id="login">
                    <strong>Bonjour, <span class="username"><%@ Page.User.Identity.Name %></span></strong>
                    | <span id="header-account-clock">00:00:00</span>
                    | <span id="header-account-date">-------</span>
                </section>
            </div>
        </div>
        <div id="nav-tabs">
        </div>
        <div id="container" class="container">
            <div id="mco-common-div" class="active-tab"></div>
            <div id="mco-ad-div" class="inactive-tab"></div>
            <div id="mco-besr-div" class="inactive-tab"></div>
            <div id="mco-app-div" class="inactive-tab"></div>
            <div id="mco-space-div" class="inactive-tab"></div>
        </div>
        <div id="footer">
            <span>&copy; <%@ DateTime.Now.Year %> - Structis Maroc<%@ Html.ActionLink("mco", "Index", null, new { @class = "footer-logo-img" })%>
            - MCO Easy Tool : Version 3.1.07
            </span>
        </div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="ScriptsSection" runat="server">
    <script src="/Scripts/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="/Scripts/modernizr-2.5.3.js" type="text/javascript"></script>
    <script src="/Scripts/jquery-impromptu.js" type="text/javascript"></script>
    <script type="text/javascript">
        SITENAME = "/";

        var loadinggif = "<div class='loading-gif-div'>" +
                    "<div class='transparent'></div>" +
                    "<div class='anti-transparent'><div class='loading-gif-img'>" +
                    "<span class='waiting-span'>Veuillez patienter s'il vous plait....</span>" +
                    "<span class='server-msg'>....<br />" +
                    "<img src='Images/mini-loading.gif' style='position:relative;'/>" +
                    "</div></div></span></div>";

        var APP_MCO_DIV = $("#mco-app-div");
        var BESR_MCO_DIV = $("#mco-besr-div");
        var AD_MCO_DIV = $("#mco-ad-div");
        var SPACE_MCO_DIV = $("#mco-space-div");
        var COMMON_MCO_DIV = $("#mco-common-div");


        var NAV_HOME_BTN = $("#nav-home-btn");
        var NAV_CONTACT_BTN = $("#nav-contact-btn");
        var NAV_TOOLS_BTN = $("#nav-tools-btn");
        var NAV_ACCOUNTS_BTN = $("#nav-accounts-btn");

        var NAV_AD_SCAN_SCHEDULER_BTN = $("#scan-scheduler");
        var NAV_AD_REPORTS_BTN = $("#nav-reports-btn");
        var NAV_AD_FAULTYSERVERS_BTN = $("#nav-faultyservers-btn");
        var NAV_AD_SETTINGS_BTN = $("#nav-adsettings-btn");
        var NAV_AD_RECIPIENTS_BTN = $("#nav-recipients-btn");

        var NAV_BESR_CHECKS_INIT_BTN = $("#nav-besrreportsinit-btn");
        var NAV_BESR_REPORTS_BTN = $("#nav-besrreports-btn");
        var NAV_BESR_SCHEDULER_BTN = $("#besr-scheduler");
        var NAV_BESR_POOLS_BTN = $("#nav-pools-btn");
        var NAV_BESR_FAILED_BTN = $("#failed-backupservers-btn");
        var NAV_BESR_INIT_POOLS_BTN = $("#init-pools-btn");
        var NAV_BESR_RECIPIENTS_BTN = $("#nav-besrrecipients-btn");

        var NAV_APP_CHECKS_INIT_BTN = $("#nav-appreportsinit-btn");
        var NAV_APP_APPLICATIONS_BTN = $("#nav-applications-btn");
        var NAV_APP_SCHEDULER_BTN = $("#app-scheduler");
        var NAV_APP_REPORTS_BTN = $("#nav-appreports-btn");
        var NAV_APP_FAILED_BTN = $("#failed-applications-btn");
        var NAV_APP_INIT_APPLICATIONS_BTN = $("#init-applications-btn");
        var NAV_APP_RECIPIENTS_BTN = $("#nav-apprecipients-btn");
        var NAV_APP_DOMAINS_BTN = $("#nav-appdomains-btn");
        var NAV_APP_FURTHER_CHECKS_INIT_BTN = $("#nav-detailledappreportsinit-btn");

        var NAV_SPACE_CHECKS_INIT_BTN = $("#nav-spacereportsinit-btn");
        var NAV_SPACE_REPORTS_BTN = $("#nav-spacereports-btn");
        var NAV_SPACE_SCHEDULER_BTN = $("#space-scheduler");
        var NAV_SPACE_SERVERS_BTN = $("#nav-spaces-btn");
        var NAV_SPACE_FAILED_BTN = $("#failed-spaceservers-btn");
        var NAV_SPACE_INIT_SERVERS_BTN = $("#init-spaceservers-btn");
        var NAV_SPACE_RECIPIENTS_BTN = $("#nav-spacerecipients-btn");

    </script>
</asp:Content>
