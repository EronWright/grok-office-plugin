/// <reference path="../App.js" />
/// <reference path="Streams.js" />

(function () {
    "use strict";

    var manager = null;
    
    // the data source will be constructed by the form
    var dataSource = null;

    var binding = null;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize({}, function () {
                manager = new Sample.StreamBindingFactory();

                $('#stream-create-binding-from-selection').click(createBindingFromSelection);
                $('#stream-create-cancel').click(cancel);
                $('#stream-create-submit').click(submit);
                $('#stream-datasource-list').accordion({
                    collapsible: true
                });
                $('#stream-datasource-panel').hide();

                initialize();
            });
        });
    };

    function initialize() {

        // the order of operations is:
        // 1. create the binding from the selection (which allows all rows to be encompassed without requiring the user to select the rows)
        // 2. inspect the binding to infer the schema spec
        // 3. present a UI to confirm stream spec
        // 4. create the stream in Grok
        // 5. sync the stream
    }

    function createBindingFromSelection() {

        var bindingId = app.generateGuid();

        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: bindingId },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Create Stream', asyncResult.error.message);
                } else {
                    binding = asyncResult.value;

                    // infer the data source binding
                    manager.createDataSourceFromBindingAsync(binding, function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            Office.context.document.bindings.releaseByIdAsync(binding.id, function () {
                                app.showNotification('Create Stream', asyncResult.error.message);
                            });
                            return;
                        }

                        dataSource = asyncResult.value;

                        // populate the UI with datasource information.
                        var fieldNames = $.map(dataSource.fields, function (field) {
                            return field.name;
                        })

                        $('<h3 id="datasource-header-' + dataSource.name + '">' + 'Table 1' + '</h3>').appendTo('#stream-datasource-list');
                        $('<div id="datasource-data-' + dataSource.name + '">' + fieldNames.join(', ') + '</div>').appendTo('#stream-datasource-list');

                        $('#stream-datasource-select').hide();
                        $('#stream-datasource-panel').show();
                        $('#stream-datasource-list').accordion("refresh");
                    });
                }
            }
        );

    }

    function submit() {
        
        var streamDef = {
            name: $("#stream-create-name").val(),
            dataSources: [
                dataSource
            ]
        };

        // validation
        if (!validate(streamDef)) {
            app.showNotification('Create Stream', 'Please fill in the required information.');
            return false;
        }

        // create the stream
        app.project.createStream(streamDef, function (err, stream) {

            if (err) {
                app.showNotification('Create Stream', err);
                return false;
            }

            // persist the binding
            manager.bindAsync(binding, stream, function (asyncResult) {
                
                // return to the home page
                window.navigate("../Home/Home.html");
            });
        });

        return false;
    }

    function validate(streamDef) {
        if (streamDef.name.length < 1) return false;
        if (streamDef.dataSources.length < 1) return false;
        for (var i = 0; i < streamDef.dataSources.length; i++)
            if (streamDef.dataSources[i].fields.length < 1) return false;
        return true;
    }

    function cancel() {

        // delete the nascent binding
        if (binding) {
            Office.context.document.bindings.releaseByIdAsync(binding.id, function (asyncResult) {
                window.navigate("../Home/Home.html");
            });
        }
        else {
            window.navigate("../Home/Home.html");
        }
    }


})();