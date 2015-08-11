/// <reference path="../App.js" />
/// <reference path="Models.js" />

(function () {
    "use strict";

    var stream = null;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize({}, function (asyncResult) {

                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Error', asyncResult.error.message);
                    return;
                }

                var streamId = $.getQueryString('streamId');
                app.client.getStream(streamId, function (err, s) {
                    if (err) {
                        app.showNotification('Error', err);
                        return;
                    }

                    stream = s;

                    // populate predicted field list
                    $('#model-stream-name').text(stream.get('name'));

                    var predictedFieldList = $('#model-create-predictedField')[0];
                    var fields = stream.get("dataSources")[0].fields;
                    $.each(fields, function (i, field) {
                        predictedFieldList[predictedFieldList.length] = new Option(field.name, field.name); 
                    }); 

                });

            });

            $('#model-create-cancel').click(cancel);
            $('#model-create-submit').click(submit);

        });
    };

    function submit() {

        var modelDef = {
            name: $("#model-create-name").val(),
            streamId: stream.getId(),
            predictedField: $('#model-create-predictedField').val(),
        };

        if (modelDef.name.length < 1) return false;

        // create the model
        app.project.createModel(modelDef, function (err, model) {

            if (err) {
                app.showNotification('Error', err);
                return;
            }

            // create a model binding
            var factory = new Sample.ModelBindingFactory();
            factory.createFromModelAsync(model, $.proxy(function (asyncResult) {

                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Create Model', 'Table Creation failed:' + asyncResult.error.message);
                    return;
                }

                window.navigate("../Home/Home.html");

            }, this));

            // return to the home page
            
        });

        return false;
    }

    function cancel() {

        // return to the home page
        window.navigate("../Home/Home.html");
    }

})();