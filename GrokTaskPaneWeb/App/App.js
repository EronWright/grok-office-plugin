/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};

    app.AsyncResult = function () {
        this.asyncContext = null;
        this.status = Office.AsyncResultStatus.Succeeded;
        this.error = null;
        this.value = null;
    }

    // Common initialization function (to be called from each page)
    app.initialize = function (opts, callback) {
        $('body').append(
            '<div id="notification-message" class="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close" class="notification-message-close"></div>' +
                    '<div id="notification-message-header" class="notification-message-header"></div>' +
                    '<div id="notification-message-body" class="notification-message-body"></div>' +
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

        app.urlParam = function (name) {
            var results = new RegExp('[\\?&]' + name + '=([^&#]*)').exec(window.location.href);
            if (!results) { return 0; }
            return decodeURIComponent(results[1].replace(/\+/g, " ")) || 0;
        }

        initializeGrokAsync(opts, callback || function (asyncResult) { });
    };



    function initializeGrokAsync(opts, callback) {

        var apiKey = sessionStorage.getItem('APIKEY');
        if (!apiKey) {
            if (!opts.anonymous) {
                window.navigate("/Home/Login.html");
                return;
            }
            var ar = new app.AsyncResult();
            callback(ar);
            return;
        }

        var client = new GROK.Client(
            apiKey, {
                //endpoint: 'https://api.numenta.com/',
                proxyEndpoint: '/_grok',
            });

        client.init(function (err) {
            if (err) {
                var ar = new app.AsyncResult();
                ar.status = Office.AsyncResultStatus.Failed;
                ar.error = new Error("Unable to initialize the Grok client.  " + err);
                callback(ar);
                return;
            }

            app.client = client;

            var projectRef = Office.context.document.settings.get("Project");
            if (!projectRef || !projectRef.id) {
                if (!opts.noproject) {
                    window.navigate("/App/Home/Project.html");
                    return;
                }
                var ar = new app.AsyncResult();
                callback(ar);
                return;
            }

            app.client.getProject(projectRef.id, $.proxy(function (err, project) {
                if (err) {
                    Office.context.document.settings.set("Project", null);
                    window.navigate("/App/Home/Project.html?err=" + encodeURIComponent(String(err)));
                    return;
                }

                app.project = project;
                var ar = new app.AsyncResult();
                callback(ar);
            }, this));
        });
    }

    app.parseGrokDate = function (date) {
        //assumes that dates are in UTC
        var value = new Date(date.replace(/\s/g, "T") + "Z");
        return value;
    }

    app.parseExcelDate2 = function(serialDateNumber) {
        var ExcelEpoch = new Date(1900, 0, 1);
        var jsEpoch = new Date();
        jsEpoch.setTime(0);

        if (serialDateNumber >= 61) serialDateNumber--;

        var ms = ((serialDateNumber - 1) * 24 * 60 * 60 * 1000) - (jsEpoch.getTime() - ExcelEpoch.getTime());
        var d = new Date();
        d.setTime(ms);

        return d;
    }


    app.parseExcelDate = (function () {
        var epoch = new Date(1899, 11, 30);
        var msPerDay = 8.64e7;

        return function (oaDate) {

            var days = Math.floor(oaDate);
            var offsetMs = (Math.abs(oaDate) * msPerDay) - (Math.abs(Math.floor(oaDate)) * msPerDay);

            var date = new Date(epoch.getTime());
            date.setDate(date.getDate() + days);
            date.setTime(date.getTime() + offsetMs);
            return date;
        }
    }());


    app.generateGuid = function () {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }

    return app;
})();

; (function ($) {

    $.widget("sample.confirm", {

        options: {
            title: "Confirm",
            text: "",
            buttonText: {
                ok: "OK",
                cancel: "Cancel"
            },
            onSelection: null,
        },

        _create: function () {
            this.element.addClass("sample-confirm");
            this.element.addClass("notification-message");

            $(
                '<div class="padding">' +
                    '<div class="notification-message-close"></div>' +
                    '<div class="notification-message-header"></div>' +
                    '<div class="notification-message-body"></div>' +
                    '<div class="notification-message-buttons"><button class="ok-button"/><button class="cancel-button"/></div>' +
                '</div>'
            ).appendTo(this.element);

            $(".notification-message-close", this.element)
                .click($.proxy(this._onCancel, this));

            $(".cancel-button", this.element)
                .button()
                .click($.proxy(this._onCancel, this));

            $(".ok-button", this.element)
                .button()
                .click($.proxy(this._onOK, this));

            this._refresh();
        },

        show: function (onSelection) {
            if (onSelection) {
                this.options.onSelection = onSelection;
            }
            this.element.slideDown('fast');
        },

        _refresh: function () {
            $('.notification-message-header', this.element).text(this.options.title);
            $('.notification-message-body', this.element).text(this.options.text);
            $(".ok-button", this.element).button({ label: this.options.buttonText.ok });
            $(".cancel-button", this.element).button({ label: this.options.buttonText.cancel });
        },

        _onOK: function () {
            this.element.hide();
            if (this.options.onSelection) {
                this.options.onSelection(true);
            }
            return false;
        },

        _onCancel: function () {
            this.element.hide();
            if (this.options.onSelection) {
                this.options.onSelection(false);
            }
            return false;
        },

        _destroy: function () {
            this.bindButton.remove();
            this.element.removeClass("notification-message");
            this.element.removeClass("sample-confirm");
        },

        _setOptions: function () {
            this._superApply(arguments);
            this._refresh();
        },

        //_setOption: function( key, value ) {
        //    this._super( key, value );
        //}
    });

    $.extend({
        getQueryString: function (name) {
            function parseParams() {
                var params = {},
                    e,
                    a = /\+/g,  // Regex for replacing addition symbol with a space
                    r = /([^&=]+)=?([^&]*)/g,
                    d = function (s) { return decodeURIComponent(s.replace(a, " ")); },
                    q = window.location.search.substring(1);

                while (e = r.exec(q))
                    params[d(e[1])] = d(e[2]);

                return params;
            }

            if (!this.queryStringParams)
                this.queryStringParams = parseParams();

            return this.queryStringParams[name];
        }
    });
})(jQuery);


(function (global) {
    "use strict";

    global.Sample = {};

    global.Sample.AsyncResult = function () {
        this.asyncContext = null;
        this.status = Office.AsyncResultStatus.Succeeded;
        this.error = null;
        this.value = null;
    }
})(window);