var app = (function () {
    "use strict";

    var app = {};

    app._renewStates = [];

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        var authContext = new AuthenticationContext({
            tenant: 'common',
            clientId: '749e5c27-c434-4267-bf35-8a863013e783',
            postLogoutRedirectUri: window.location.href.split("?")[0].split("#")[0],
            redirectUri: window.location.href.split("?")[0].split("#")[0],
            endpoints: {
                'https://graph.microsoft.com': 'https://graph.microsoft.com',
            },
            cacheLocation: 'localStorage'
        });

        var $userDisplay = $(".app-user");
        var $signInButton = $(".app-login");
        var $signOutButton = $(".app-logout");

        var isCallback = authContext.isCallback(window.location.hash);
        authContext.handleWindowCallback();

        if (isCallback && !authContext.getLoginError()) {
            authContext.info("isCallback");
            window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
        }

        var user = authContext.getCachedUser();
        if (user) {
            $userDisplay.html(user.userName);
            $userDisplay.show();
            $signInButton.hide();
            $signOutButton.show();
        } else {
            $userDisplay.empty();
            $userDisplay.hide();
            $signInButton.show();
            $signOutButton.hide();
        }

        $signOutButton.click(function () {
            authContext.logOut();
        });
        $signInButton.click(function () {
            authContext.login();
        });

        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });

        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };

        app.acquireToken = function (adal, resource, callback) {
            adal.info("app.acquireToken for " + resource);
            adal.acquireToken(resource, function (error, token) {
                if (error || !token) {
                    adal.info("app.acquireToken performing login");
                    app.login(adal, resource, function (error, token) {
                        callback(null, token);
                    });
                }
                else {
                    callback(null, token);
                };
            });
        }

        app.login = function (adal, resource, callback) {
            adal.config.displayCall = function (urlNavigate) {
                app._userFrame(adal, urlNavigate);
            };
            adal.info("start adal.login()");
            adal.login();
            adal.info("adal.login() finished");
            adal.config.displayCall = null;
            //adal.acquireToken(resource, callback);
        }

        /**
         * Redirects a visible IFRAME to Azure AD authorization endpoint.
         * @param {object}   adal  authorization context.
         * @param {string}   urlNavigate  Url of the authorization endpoint.
         */
        app._userFrame = function (adal, urlNavigate) {
            var userFrame = app._addUserFrame(adal._guid());
            urlNavigate = adal._addHintParameters(urlNavigate);
            urlNavigate += "&iframe=true";
            adal.info("userFrame.src = " + urlNavigate);
            userFrame.location.replace(urlNavigate);

            var registeredRedirectUri = "";

            if (adal.config.redirectUri.indexOf('#') != -1) {
                registeredRedirectUri = adal.config.redirectUri.split("#")[0];
            }
            else {
                registeredRedirectUri = adal.config.redirectUri;
            }

            adal.info("start polling for " + registeredRedirectUri);
            var pollTimer = window.setInterval(function () {
                adal.info("polling...");
                try {
                    if (userFrame.src.indexOf(registeredRedirectUri) != -1) {
                        adal.handleWindowCallback(userFrame.src.hash);
                        window.clearInterval(pollTimer);
                        adal._loginInProgress = false;
                        adal.info("Closing userFrame");
                        userFrame.parentNode.removeChild(userFrame);
                    }
                } catch (e) {
                    adal.error("Error: " + JSON.stringify(e));
                }
            }, 1000);
            adal.info("finished polling.");
        };

        // Creates a full size IFRAME embedded in the current window
        app._addUserFrame = function (iframeId) {
            if (document.createElement && document.documentElement &&
                (window.opera || window.navigator.userAgent.indexOf('MSIE 5.0') === -1)) {
                var ifr = document.createElement('iframe');
                ifr.setAttribute('id', iframeId);
                ifr.style.position = 'absolute';
                ifr.style.left = ifr.style.top = ifr.style.right = ifr.style.bottom = ifr.borderWidth = '0';
                ifr.style.width = ifr.style.height = '80%'; // TODO change to 100%
                ifr.style.zIndex = 10;
                var adalFrame = document.getElementsByTagName('body')[0].appendChild(ifr);
            }
            else if (document.body && document.body.insertAdjacentHTML) {
                document.body.insertAdjacentHTML('beforeEnd', '<iframe name="' + iframeId + '" id="' + iframeId + '" style="left:0; top:0; right:0; bottom:0; border-width:0;"></iframe>');
            }
            if (window.frames && window.frames[iframeId]) {
                return window.frames[iframeId];
            }
        }
    };

    return app;
})();