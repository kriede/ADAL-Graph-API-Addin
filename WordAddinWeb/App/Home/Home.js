/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            if (isAdmin) {
                $('#admin-consent').click(adminConsent)

            } else {
                $('#admin-consent').hide();
            }
            $('#use-graph-api').click(useGraphAPI);
        });
    };

    // Request admin to consent with requested permissions
    // https://blog.mastykarz.nl/implementing-admin-consent-multitenant-office-365-applications-implicit-oauth-flow/
    function adminConsent() {
        var adal = new AuthenticationContext();
        adal.config.displayCall = function adminFlowDisplayCall(urlNavigate) {
            urlNavigate += '&prompt=admin_consent';
            adal.promptUser(urlNavigate);
        };
        adal.login();
        adal.config.displayCall = null;
    }

    // Reads email address from current document selection and displays user information
    function useGraphAPI() {
        var baseEndpoint = 'https://graph.microsoft.com';
        var authContext = new AuthenticationContext(config);
        var result = $("#results");
        result.html("use graph api");

        authContext.acquireToken(baseEndpoint, function (error, token) {
            if (error || !token) {
                app.showNotification("No token: " + error);
                return;
            }
            result.html("got token...");
            var email = authContext._user.userName;
            var url = "https://graph.microsoft.com/v1.0/me/";
            var html = "<ul>";
            $.ajax({
                beforeSend: function (request) {
                    result.html("before send...");
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: url,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + token,
                }
            }).done(function (response) {
                result.html("request done...");
                html += getPropertyHtml("Display name", response.displayName);
                html += getPropertyHtml("Address", response.streetAddress);
                html += getPropertyHtml("Postal code", response.postalCode);
                html += getPropertyHtml("City", response.city);
                html += getPropertyHtml("Country", response.country);
                html += getPropertyHtml("Photo", response.thumbnailPhoto);
                $("#results").html(html);
                app.showNotification(html);
            }).fail(function (response) {
                app.showNotification(response.responseText);
            });
        });
    }

    function getPropertyHtml(key, value) {
        return "<li><strong>" + key + "</strong> : " + value + "</li>";
    }

    // https://blog.mastykarz.nl/implementing-admin-consent-multitenant-office-365-applications-implicit-oauth-flow/
    function isAdmin() {
        var deferred = $q.defer();

        $http({
            url: 'https://graph.windows.net/me/memberOf?api-version=1.6',
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=nometadata'
            }
        }).success(function (data) {
            var isAdmin = false;

            for (var i = 0; i < data.value.length; i++) {
                var obj = data.value[i];

                if (obj.objectType === 'Role' &&
                  obj.isSystem === true &&
                  obj.displayName === 'Company Administrator') {
                    isAdmin = true;
                    break;
                }
            }

            deferred.resolve(isAdmin);
        }).error(function (err) {
            deferred.reject(err);
        });

        return deferred.promise;
    }

})();
