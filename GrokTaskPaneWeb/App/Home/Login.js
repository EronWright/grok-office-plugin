/// <reference path="../App.js" />

(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize({
                anonymous: true,
            },
            $.proxy(function () {

                $('#progressbar').progressbar({
                    value: false
                });

                $('#login-submit').click(function () {

                    $('#login').hide();
                    $('#progress-panel').show();

                    var apiKey = $("#login-apikey").val();
                    if (apiKey.length < 1) return;

                    var client = new GROK.Client(
                        apiKey, {
                            //endpoint: 'https://api.numenta.com/',
                            proxyEndpoint: '/_grok'
                        });


                    client.init(function (err) {

                        if (err) {
                            $('#progress-panel').hide();
                            $('#login').show();

                            app.showNotification('Sign In', 'Unable to connect to Grok.  ' + err);
                            return;
                        }

                        window.sessionStorage.setItem("APIKEY", apiKey);

                        var projectRef = Office.context.document.settings.get("Project");
                        if (!projectRef) {
                            window.navigate("Project.html");
                        } else {
                            window.navigate("Home.html");
                        }
                    });

                    return false;
                });
            }, this));
        });
    };

})();