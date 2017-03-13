/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#get-data-from-selection').click(useGraphAPI);
        });
    };

    // Reads email address from current document selection and displays user information
    function useGraphAPI() {
        var baseEndpoint = 'https://graph.microsoft.com';
        var authContext = new AuthenticationContext(config);

        authContext.acquireToken(baseEndpoint, function (error, token) {
            if (error || !token) {
                app.showNotification("No token: " + error);
                return;
            }
            var email = authContext._user.userName;
            var url = "https://graph.microsoft.com/v1.0/me/";
            var html = "<ul>";
            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: url,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + token,
                }
            }).done(function (response) {
                html += getPropertyHtml("Display name", response.displayName);
                html += getPropertyHtml("Address", response.streetAddress);
                html += getPropertyHtml("Postal code", response.postalCode);
                html += getPropertyHtml("City", response.city);
                html += getPropertyHtml("Country", response.country);
                html += getPropertyHtml("Photo", response.thumbnailPhoto);
                $("#results").html(html);
            }).fail(function (response) {
                app.showNotification(response.responseText);
            });
        });
    }

    function getPropertyHtml(key, value) {
        return "<li><strong>" + key + "</strong> : " + value + "</li>";
    }

})();
