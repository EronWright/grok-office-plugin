/// <reference path="../App.js" />
/// <reference path="Streams/Streams.js" />
/// <reference path="Models/Models.js" />

(function () {
    "use strict";

    var factory;
    var modelFactory;
    var streams = [];
    var models = [];

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize({}, function () {

                $('#project-name').text(app.project.getName()).hide();
                $("#project-menu-list").menu({
                    select: _onProjectMenuSelected
                }).hide();
                $("#project-menu-button").button({
                    label: app.project.getName(),
                    icons: {
                        secondary: "ui-icon-triangle-1-s"
                    },
                }).click(function () {
                    var menu = $("#project-menu-list").show().position({
                        my: "right top",
                        at: "right bottom",
                        of: this
                    });
                    $(document).one("click", function () {
                        menu.hide();
                    });
                    return false;
                });

                $("#tabs").tabs({
                    activate: function (event, ui) {
                        $('.ui-accordion', ui.newPanel).accordion("refresh");
                    },
                });

                $('#createStream').click(createStream);

                $("#stream-list").accordion();
                $("#model-list").accordion();

                factory = new Sample.StreamBindingFactory();
                modelFactory = new Sample.ModelBindingFactory();

                app.project.listStreams(function (err, s) {
                    streams = s;
                    $.each(streams, function (i, stream) {
                        var streamId = stream.getId();
                        var header =
                            $("#stream-header-template").clone().attr("id", 'stream-header-' + streamId).show()
                            .appendTo('#stream-list');
                        $(".stream-header-name", header).text(stream.getScalar("name"));

                        var div =
                            $("#stream-item-template").clone().attr("id", 'stream-data-' + streamId).show()
                            .appendTo('#stream-list').stream({ streamId: streamId });
                    });

                    $("#stream-list").accordion("refresh");

                    app.project.listModels(function (err, s) {
                        models = s;
                        $.each(models, function (i, model) {
                            var modelId = model.getId();

                            var header =
                                $("#model-header-template").clone().attr("id", 'model-header-' + modelId).show()
                                .appendTo('#model-list');
                                $(".model-header-name", header).text(model.getScalar("name"));

                            var div =
                                $("#model-item-template").clone().attr("id", 'model-data-' + modelId).show()
                                .appendTo('#model-list').model({ modelId: modelId });
                        });

                        $("#model-list").accordion("refresh");
                    });
                });
            });
        });
    };

    function _onProjectMenuSelected(event, item) {

        switch (event.srcElement.hash) {
            case "#change":
                _changeProject();
                break;
            case "#delete":
                $('#confirm-dialog').confirm({
                    title: "Confirmation",
                    text: "Are you sure you want to delete the project '" + app.project.getName() + "'?",
                    buttonText: {
                        ok: "Delete",
                        cancel: "Cancel"
                    },
                }).confirm("show", $.proxy(function (success) {
                    if (!success) return;
                    app.project.delete(function (err) {
                        if (err) {
                            app.showNotification('Delete Project', err);
                            return;
                        }
                        _changeProject();
                    });
                }));
                break;
        }
    }

    function _changeProject() {
        Office.context.document.settings.set("Project", null);
        Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
            window.navigate("/App/Home/Project.html");
        });
    }

    function createStream() {
        window.navigate("../Streams/CreateStream.html");
    }

    $.widget( "sample.stream", {

        options: {
            streamId: "",
        },

        _create: function() {
            this.element.addClass("sample-stream");

            this.unboundPanel = $(".stream-unbound-panel", this.element);
            this.boundPanel = $(".stream-bound-panel", this.element);

            this.menuButton = $(".menuButton", this.element).button({
                icons: {
                    primary: "ui-icon-gear",
                    secondary: "ui-icon-triangle-1-s"
                },
                text: false
            }).click(function () {
                var menu = $(".menuList", $(this).parent()).show().position({
                    my: "right top",
                    at: "right bottom",
                    of: this
                });
                $(document).one("click", function () {
                    menu.hide();
                });
                return false;
            });

            this.menuList = $(".menuList", this.element).menu().hide();

            this._on(this.menuList, {
                menuselect: "select"
            });

            this.pasteButton = $(".stream-paste-button", this.element).button({
                icons: {
                    primary: "ui-icon-clipboard"
                }
            }).click($.proxy(this._onPasteButtonClicked, this));

            this.syncButton = $(".stream-sync-button", this.element).button({
                icons: {
                    primary: "ui-icon-transferthick-e-w",
                },
            }).click($.proxy(this._onSyncButtonClicked, this));

            this.uploadProgressBar = $(".stream-progressbar", this.element).progressbar({
                value: 0,
                max: 0,
                disabled : true
            });
            
            $('.stream-model-button', this.element).button({
                icons: {
                    primary: "ui-icon-shuffle",
                },
            }).click($.proxy(this.createModel, this));

            this._refresh();
        },

        _refresh: function () {

            this.stream = streams.filter(function (s) { return s.getId() == this.options.streamId }, this)[0];
            this.binding = null;
            factory.getBindingForStreamAsync(this.stream, $.proxy(function (asyncResult) {
                var tb = asyncResult.value;
                this.binding = tb;

                if (!this.binding) {
                    this.boundPanel.hide();
                    this.unboundPanel.show();
                }
                else {
                    this.unboundPanel.hide();
                    this.boundPanel.show();
                }
                if (!tb) {
                    this.boundPanel.hide();
                    this.unboundPanel.show();
                }
                else {
                    this.binding = tb;

                    this.unboundPanel.hide();
                    this.boundPanel.show();

                    tb.setUploadListener($.proxy(function (uploadState) {
                        switch (uploadState.status) {
                            case 'stopped':
                                this.syncButton.button({ disabled: false });
                                this.uploadProgressBar.progressbar({
                                    value: 0,
                                    max: 0,
                                    disabled: true
                                });
                                break;

                            case 'running':
                                this.syncButton.button({ disabled: true });
                                this.uploadProgressBar.progressbar({
                                    value: uploadState.uploadedCount,
                                    max: uploadState.rowCount,
                                    disabled: false
                                });
                                break;
                        }
                    }, this));

                    tb.binding.addHandlerAsync(Office.EventType.BindingSelectionChanged, $.proxy(this._onBindingSelectionChanged, this));
                }
            }, this));
        },

        _onBindingSelectionChanged: function (eventArgs) {
            var lastId = window.sessionStorage.getItem("LASTBINDINGID");
            if (lastId === eventArgs.binding.id) return;
            window.sessionStorage.setItem("LASTBINDINGID", eventArgs.binding.id);

            $('#tabs ul:first li:eq(0) a').click();
            $('#stream-header-' + this.options.streamId).click();
        },

        _onPasteButtonClicked: function(eventArgs) {
            this.download();
        },

        _onSyncButtonClicked: function (eventArgs) {
            this.sync();
        },

        // binds an existing stream to a new Excel table
        download: function () {

            factory.createFromStreamAsync(this.stream, $.proxy(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Stream', 'Download failed:' + asyncResult.error.message);
                    return;
                }

                this.binding = asyncResult.value;
                this._refresh();
            }, this));
        },

        // uploads stream data
        sync: function () {

            this.binding.uploadAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Stream', 'Sync failed:' + asyncResult.error.message);
                }
                else {
                    app.showNotification('Stream', 'Synchronization is complete.');
                }
            });
        },

        select: function (event, item) {
            switch (event.srcElement.hash) {
                case "#delete":
                    var streamId = this.stream.getId();

                    if (models.filter(function (m) { return m.get('streamId') === streamId }, this).length >= 1) {
                        app.showNotification('Error', 'Unable to delete a stream that is in use by a model.');
                        return;
                    }
                    
                    $('#confirm-dialog').confirm({
                        title: "Confirmation",
                        text: "Are you sure you want to delete the stream '" + this.stream.get("name") + "'?",
                        buttonText: {
                            ok: "Delete",
                            cancel: "Cancel"
                        },
                    }).confirm("show", $.proxy(function (success) {
                        if (!success) return;
                        // release binding

                        this.stream.delete($.proxy(function (err) {

                            if (err) {
                                app.showNotification('Delete Stream', err);
                                return;
                            }

                            // note that stream object is unusable now

                            streams = streams.filter(function (s) { return (s._scalars); });

                            // remove from UI
                            $("#stream-header-" + streamId).remove();
                            $("#stream-data-" + streamId).remove();
                            $("#stream-list").accordion("refresh");

                            if (this.binding) {
                                factory.releaseBindingAsync(this.binding, streamId, $.proxy(function (asyncResult) {
                                    // deletion is complete
                                }, this));
                            }
                            else {
                                // deletion is complete
                            }

                        }, this));
                    }, this));

                    break;

                case "#create-model":
                    this.createModel();
                    break;
            }

        },

        _destroy: function() {
            this.element.removeClass("sample-stream");
        },

        _setOptions: function() {
            this._superApply( arguments );
            this._refresh();
        },

        //_setOption: function( key, value ) {
        //    this._super( key, value );
        //}

        createModel: function () {
            window.navigate("../Models/CreateModel.html?streamId=" + encodeURIComponent(this.options.streamId));
        }
    });

    $.widget("sample.model", {

        options: {
            modelId: "",
        },

        _create: function () {
            var me = this;

            this.element
                .addClass("sample-model");

            this.unboundPanel = $(".model-unbound-panel", this.element);
            this.boundPanel = $(".model-bound-panel", this.element);
            
            this.pasteButton = $(".model-paste-button", this.element).button({
                icons: {
                    primary: "ui-icon-clipboard"
                }
            }).click($.proxy(this.download, this));

            this.menuButton = $(".menuButton", this.element).button({
                icons: {
                    primary: "ui-icon-gear",
                    secondary: "ui-icon-triangle-1-s"
                },
                text: false
            }).click(function () {
                var menu = $(".menuList", $(this).parent()).show().position({
                    my: "right top",
                    at: "right bottom",
                    of: this
                });
                $(document).one("click", function () {
                    menu.hide();
                });
                return false;
            });

            this.menuList = $(".menuList", this.element).menu().hide();
            this._on(this.menuList, {
                menuselect: "_select"
            });

            this.swarmPanel = $(".model-swarm-panel", this.element);

            this.swarmMenu = $(".model-swarm-menu", this.element).menu().hide();
            this._on(this.swarmMenu, { menuselect: "_onSwarmSelected" });
            this.swarmButton = $(".model-swarm-button", this.element).button({
                icons: {
                    primary: "ui-icon-search",
                    secondary: "ui-icon-triangle-1-s"
                },
            }).click(function () {
                me.swarmMenu.show().position({ my: "left top", at: "left bottom", of: this });
                $(document).one("click", function() {
                    me.swarmMenu.hide();
                });
                return false;
            });
            this.swarmProgressBar = $(".model-swarm-progressbar", this.element).progressbar({
                value: 0,
                max: 0,
                disabled: true
            });
            
            this.promotePanel = $(".model-promote-panel", this.element);
            this.promoteButton = $(".model-promote-button", this.element).button({
                icons: {
                    primary: "ui-icon-check",
                },
            }).click($.proxy(this.promote, this));
            this.promoteCompletedCell = $(".model-swarm-completed", this.element);
            this.promoteFieldsUsedCell = $(".model-swarm-fields-used", this.element);
            this.promoteAverageErrorCell = $(".model-swarm-average-error", this.element);

            this.runningPanel = $(".model-running-panel", this.element);

            this.promoteButton = $(".model-stop-button", this.element).button({
                icons: {
                    primary: "ui-icon-stop",
                },
            }).click($.proxy(this.stopModel, this));

            // initialize data and refresh
            //this.model = models.filter(function (o) { return o.getId() == this.options.modelId }, this)[0];
            app.project.getModel(this.options.modelId, $.proxy(function (err, model) {
                this.model = model;
                this.stream = streams.filter(function (s) { return s.getId() == this.model.get('streamId') }, this)[0];

                modelFactory.getBindingForModelAsync(this.model.getId(), $.proxy(function (asyncResult) {
                    var tb = asyncResult.value;
                    if (!tb) {
                        this.binding = null;
                    }
                    else {
                        this._setBinding(new Sample.TableBindingToModel(modelFactory, tb, this.model));
                    }
                    this._refresh();
                }, this));
            }, this));
        },

        _setBinding: function (binding) {
            this.binding = binding;
            this.binding.onModelStatusChange($.proxy(this._onModelStatusChanged, this));
            this.binding.binding.addHandlerAsync(Office.EventType.BindingSelectionChanged, $.proxy(this._onBindingSelectionChanged, this));
            this.binding.startAsync();

            //this.binding.monitorPredictionsAsync({
            //    onError: function (err) {
            //        app.showNotification('Model', 'Unable to download prediction results. ' + err);
            //    },
            //});
        },

        _refresh: function () {

            $('.model-stream-name', this.element).text(this.stream.get('name'));
            $('.model-predicted-field', this.element).text(this.model.getScalar('predictedField'));

            if (!this.binding) {
                this.boundPanel.hide();
                this.unboundPanel.show();
            }
            else {
                this.unboundPanel.hide();
                this.boundPanel.show();

                var persistedModel = modelFactory.persistedModels[this.model.getId()];
                var status = this.model.get("status") || GROK.Model.STATUS.STOPPED;
                switch (status) {

                    case GROK.Model.STATUS.STOPPED:
                        // configure for stopped state
                        this.runningPanel.hide();
                        this.swarmPanel.show();
                        this.swarmProgressBar.progressbar({
                            value: 0,
                            disabled: true
                        });
                        this.swarmButton.button({
                            disabled: false,
                        });

                        var lastSwarm = this.model.get("lastSwarm");

                        if (lastSwarm && lastSwarm.status === GROK.Swarm.STATUS.COMPLETED) {
                            this.promotePanel.show();
                            this.promoteCompletedCell.text(lastSwarm.details.endTime);
                            this.promoteFieldsUsedCell.text(lastSwarm.details.fieldsUsed.join(","));
                            this.promoteAverageErrorCell.text(lastSwarm.results.averageError);

                        } else {
                            this.promotePanel.hide();
                        }
                        break;

                    case GROK.Model.STATUS.SWARMING:
                        // configure for swarming state
                        this.promotePanel.hide();
                        this.runningPanel.hide();
                        this.swarmPanel.show();
                        this.swarmProgressBar.progressbar({
                            value: false,
                            disabled: false
                        });
                        this.swarmButton.button({
                            disabled: true,
                        });
                        break;

                    case GROK.Model.STATUS.STARTING:
                    case GROK.Model.STATUS.RUNNING:
                        // configure for running state
                        this.promotePanel.hide();
                        this.swarmPanel.hide();
                        this.runningPanel.show();

                        var smape = persistedModel.SMAPE.denominator != 0 ?
                            persistedModel.SMAPE.numerator / persistedModel.SMAPE.denominator : undefined;

                        $('.model-running-average-error', this.element).text(this.model.getScalar('averageError') || "N/A");
                        $('.model-running-smape', this.element).text(smape ? String((smape * 100).toFixed(2)) + '%' : "N/A");

                        break;

                    case GROK.Model.STATUS.ERROR:
                        app.showNotification('Model', 'The model is in an error state.');
                        break;
                }
            }
        },

        _onBindingSelectionChanged: function (eventArgs) {
            var lastId = window.sessionStorage.getItem("LASTBINDINGID");
            if (lastId === eventArgs.binding.id) return;
            window.sessionStorage.setItem("LASTBINDINGID", eventArgs.binding.id);

            $('#tabs ul:first li:eq(1) a').click();
            $('#model-header-' + this.options.modelId).click();
        },

        _onModelStatusChanged: function(model) {
            this.model = model;
            this._refresh();
        },

        _select: function (event, item) {
            switch (event.srcElement.hash) {
                case "#delete":

                    $('#confirm-dialog').confirm({
                        title: "Confirmation",
                        text: "Are you sure you want to delete the model '" + this.model.get("name") + "'?",
                        buttonText: {
                            ok: "Delete",
                            cancel: "Cancel"
                        },
                    }).confirm("show", $.proxy(function (success) {
                        if (!success) return;

                        if (this.binding) {
                            this.binding.unbindAsync($.proxy(function (asyncResult) {
                                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                    app.showNotification('Delete Model', 'Unbind failed:' + asyncResult.error.message);
                                    return;
                                }
                                this._deleteModel();
                            }, this));
                        }
                        else {
                            this._deleteModel();
                        }
                    }, this));
                    break;

                default:
                    break;
            }

        },

        _deleteModel: function () {
            var modelId = this.model.getId();

            this.model.delete($.proxy(function (err) {
                if (err) {
                    app.showNotification('Delete Model', err);
                    return;
                }
                    
                // note that model is unusable now

                models = models.filter(function (m) { return (m._scalars); });

                // remove from UI
                $("#model-header-" + modelId).remove();
                $("#model-data-" + modelId).remove();
                $("#model-list").accordion("refresh");

                // deletion complete
            }, this));
        },

        _destroy: function () {
            this.element.removeClass("sample-model");
        },

        _setOptions: function () {
            this._superApply(arguments);
            this._refresh();
        },

        //_setOption: function( key, value ) {
        //    this._super( key, value );
        //}

        // binds an existing model to a new Excel table
        download: function (event) {

            modelFactory.createFromModelAsync(this.model, $.proxy(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Model', 'Download failed:' + asyncResult.error.message);
                    return;
                }

                this._setBinding(asyncResult.value);
                this._refresh();
            }, this));
        },

        _onSwarmSelected: function (event, item) {
            var swarmSize = event.srcElement.hash.replace("#", "");
            this.swarm(swarmSize);
        },

        swarm: function (swarmSize) {
            this.model.startSwarm({ size: swarmSize }, $.proxy(function (err, swarm, monitor) {
                if (err) {
                    app.showNotification('Swarm Model', err);
                    return;
                }
                // do not use the built-in swarm monitor; our model monitor subsumes it.
                monitor.stop();

                this.binding.checkStatus();
                
                //this._refresh();
            }, this));
        },

        promote: function () {
            this.model.promote($.proxy(function (err) {
                if (err) {
                    app.showNotification('Promote Model', err);
                    return;
                }

                this.binding.checkStatus();
                //this._refresh();
            }, this));
        },

        stopModel: function () {
            this.model.stop($.proxy(function (err) {
                if (err) {
                    app.showNotification('Stop Model', err);
                    return;
                }

                this._refresh();
            }, this));
        },
    });
})();