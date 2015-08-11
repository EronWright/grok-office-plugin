/// <reference path="../../Scripts/grok-all.debug.js" />
/* Common app functionality */

(function (global) {
    "use strict";

    var ANY = 'any';

    global.Sample.AsyncResult = function () {
        this.asyncContext = null;
        this.status = Office.AsyncResultStatus.Succeeded;
        this.error = null;
        this.value = null;
    }

    global.Sample.ModelBindingFactory = function () {
        this.persistedModels = Office.context.document.settings.get("Models") || {};
    }

    global.Sample.ModelBindingFactory.prototype = {

        getBindingForModelAsync: function (modelId, callback) {
            var p = this.persistedModels[modelId];
            if (!p) {
                var ar = new Sample.AsyncResult();
                ar.value = null;
                callback(ar);
                return;
            }

            var bindingId = p.bindingId;
            Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {

                var binding = null;
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    // remove stale binding
                    delete this.persistedModels[modelId];
                }
                else {
                    binding = asyncResult.value;
                }
                var ar = new Sample.AsyncResult();
                ar.value = binding;
                callback(ar);
            });
        },

        createFromModelAsync: function (model, callback) {
            
            // Build a table based on the schema of the model.
            var table = new Office.TableData();

            var streamId = model.get('streamId');
            app.client.getStream(streamId, $.proxy(function (err, stream) {

                if (err) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error(err);
                        callback(ar);
                    }
                    return;
                }

                var dataSource = stream.getScalar("dataSources")[0];

                // use the timestamp field, the field that is to be predicted, and the actual predicted value as columns
                table.headers = [
                    dataSource.fields.filter(
                        function (field) { return field.dataFormat.dataType.toUpperCase() == "DATETIME"; })[0].name,
                    model.getScalar('predictedField'),
                    "Predicted " + model.getScalar('predictedField'),
                    "Confidence",
                    ];

                // Write table.
                Office.context.document.setSelectedDataAsync(
                    table,
                    { coercionType: "table" },
                    $.proxy(function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            if (callback) {
                                var ar = new Sample.AsyncResult();
                                ar.status = Office.AsyncResultStatus.Failed;
                                ar.error = new Error("Unable to create table.  Error: " + asyncResult.error.message);
                                callback(ar);
                            }
                            return;
                        }

                        // create a binding to the table
                        Office.context.document.bindings.addFromSelectionAsync(
                            Office.BindingType.Table,
                            { id: model.getId() },
                            $.proxy(function (asyncResult) {
                                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                    if (callback) {
                                        var ar = new Sample.AsyncResult();
                                        ar.status = Office.AsyncResultStatus.Failed;
                                        ar.error = new Error("Unable to create binding.  Error: " + asyncResult.error.message);
                                        callback(ar);
                                    }
                                    return;
                                }

                                var binding = asyncResult.value;
                                var modelBinding = new Sample.TableBindingToModel(this, binding, model);

                                this.bindAsync(binding, model, $.proxy(function (asyncResult) {
                                    if (callback) {
                                        var ar = new Sample.AsyncResult();
                                        ar.status = Office.AsyncResultStatus.Succeeded;
                                        ar.value = modelBinding;
                                        callback(ar);
                                    }
                                    return;
                                }, this));
                            }, this)
                        );
                    }, this)
                );
            }, this));
        },

        bindAsync: function (binding, model, callback) {
            // finalize the binding of a table to a model
            this.persistedModels[model.getId()] = {
                bindingId: binding.id,
                ROWID: -1,
                SMAPE: {
                    n: 0,
                    numerator: 0,
                    denominator: 0,
                }
            };
            this.persistAsync(callback);
        },

        releaseBindingAsync: function (binding, modelId, callback) {
            Office.context.document.bindings.releaseByIdAsync(binding.binding.id, $.proxy(function (asyncResult) {
                delete this.persistedModels[modelId];
                this.persistAsync(callback);
            }, this));
        },

        persistAsync: function(callback) {

            Office.context.document.settings.set("Models", this.persistedModels);
            Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error("Unable to persist model bindings.  Error: " + asyncResult.error.message);
                        callback(ar);
                    }
                    return;
                }

                var ar = new Sample.AsyncResult();
                callback(ar);
            });
        }
    }

    global.Sample.TableBindingToModel = function (factory, binding, model) {
        this.factory = factory;
        this.binding = binding;
        this.model = model;
        this.monitor = new Sample.ModelMonitor(this.model, {
            interval: 1000,
        });

        this.monitor.onStatusChange($.proxy(this._monitorCallback, this));
    }

    global.Sample.TableBindingToModel.prototype = {

        onModelStatusChange: function(callback) {
            this.monitor.onStatusChange($.proxy(callback, this));
        },

        checkStatus: function () {
            this.monitor.check();
        },

        startAsync: function (callback) {

            var status = this.model.get("status");
            switch (status) {
                case GROK.Model.STATUS.RUNNING:
                    if (!this.predictionMonitor) {
                        this._startPredictionMonitor();
                    }
                    break;
                default:
                    this.monitor.start();
                    break;
            }

            if (callback) {
                var ar = new Sample.AsyncResult();
                ar.status = Office.AsyncResultStatus.Succeeded;
                callback(ar);
            }
        },

        _monitorCallback: function (model) {
            this.model = model;
            
            var status = this.model.get("status");
            switch (status) {
                case GROK.Model.STATUS.RUNNING:
                    this.monitor.stop();  // do not monitor for changes after this point
                    if (!this.predictionMonitor) {
                        this._startPredictionMonitor();
                    }
                    break;

                default:
                    if (this.predictionMonitor) {
                        this.predictionMonitor.stop();
                        this.predictionMonitor = null;
                    }
                    break;
            }
        },

        _startPredictionMonitor: function (options, callback) {

            var opts = options || {};

            var rowid = this.factory.persistedModels[this.model.getId()].ROWID;

            // TODO observe model status ('running' is when prediction data is valid)

            // start a monitor for prediction data
            this.predictionMonitor = this.model.monitorPredictions({
                interval: 1000,
                lastRowIdSeen: rowid,
                outputDataOptions: {
                    limit: 1000,
                    shift: true,
                },
                onUpdate: $.proxy(this._addData, this),
                onDone: function (err) {
                    app.showNotification('Model', 'No more prediction results.');
                },
                onError: opts.onError || function () { },
            });

            if (callback) {
                var ar = new Sample.AsyncResult();
                ar.status = Office.AsyncResultStatus.Succeeded;
                callback(ar);
            }
            return;
        },

        _addData: function (output) {
            // TODO fix blocking to non-blocking conversion
            this._addDataAsync(output, null);
        },

        _addDataAsync: function (output, callback) {
            if (output.data.length == 0) {
                if (callback) {
                    var ar = new Sample.AsyncResult();
                    ar.status = Office.AsyncResultStatus.Succeeded;
                    callback(ar);
                }
                return;
            };

            var streamId = this.model.get('streamId');
            app.client.getStream(streamId, $.proxy(function (err, stream) {

                var grokDataSource = stream.getScalar("dataSources")[0];
                var convertedData = this.convertGrokDataToTableData(grokDataSource, output);

                // add rows
                this.binding.addRowsAsync(convertedData.rows, null, $.proxy(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        console.log("upload: error adding table data", asyncResult.error);
                        if (callback) {
                            var ar = new Sample.AsyncResult();
                            ar.status = Office.AsyncResultStatus.Failed;
                            ar.error = new Error("Unable to set table data.  Error: " + asyncResult.error.message);
                            callback(ar);
                        }
                        return;
                    }

                    console.log("download complete", convertedData.rows.length);

                    // store the ROWID to enable continuation in future
                    var lastROWID = output.data[output.data.length - 2][0];
                    this.factory.persistedModels[this.model.getId()].ROWID = lastROWID;

                    // recalculate the SMAPE
                    this.factory.persistedModels[this.model.getId()].SMAPE.n += convertedData.SMAPE.n;
                    this.factory.persistedModels[this.model.getId()].SMAPE.numerator += convertedData.SMAPE.numerator;
                    this.factory.persistedModels[this.model.getId()].SMAPE.denominator += convertedData.SMAPE.denominator;

                    // persist
                    this.factory.persistAsync($.proxy(function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            if (callback) {
                                var ar = new Sample.AsyncResult();
                                ar.status = Office.AsyncResultStatus.Failed;
                                ar.error = new Error("Unable to persist .  Error: " + asyncResult.error.message);
                                callback(ar);
                            }
                            return;
                        }

                        if (callback) {
                            var ar = new Sample.AsyncResult();
                            ar.status = Office.AsyncResultStatus.Succeeded;
                            callback(ar);
                        }
                        return;

                    }, this));

                }, this));

            }, this));
        },

        unbindAsync: function (callback) {

            this.monitor.stop();

            if (this.predictionMonitor) {
                this.predictionMonitor.stop();
                this.predictionMonitor = null;
            }

            this.factory.releaseBindingAsync(this, this.model.getId(), $.proxy(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error("Unable to release the binding.  Error: " + asyncResult.error.message);
                        callback(ar);
                    }
                    return;
                }
                if (callback) {
                    var ar = new Sample.AsyncResult();
                    ar.status = Office.AsyncResultStatus.Succeeded;
                    callback(ar);
                }
                return;
            }, this));
        },

        convertGrokDataToTableData: function (dataSource, output) {

            var result = {
                rows: [],
                SMAPE: {
                    n: 0,
                    numerator: 0,
                    denominator: 0,
                }
            };

            var pFieldName = this.model.getScalar('predictedField');
            var pField = dataSource.fields.filter(
                function (field) { return field.name == pFieldName; })[0];

            for (var i = 0; i < output.data.length; i++) {
                
                if (output.data[i][0] === "") continue; // the prediction row

                var row = new Array();
                var fields = [
                    output.data[i][output.meta.timestampIndex],
                    output.data[i][output.meta.predictedFieldIndex],
                    output.data[i][output.meta.predictedFieldPredictionIndex],
                    (output.data[i][output.meta.predictedFieldPredictionIndex] === "") ? "" :
                        output.data[i][4][1][output.data[i][output.meta.predictedFieldPredictionIndex]],
                ];

                row[0] = app.parseGrokDate(fields[0]);

                // use header information 
                switch (pField.dataFormat.dataType.toUpperCase()) {
                    case "DATETIME":
                        row[1] = app.parseGrokDate(fields[1]);
                        row[2] = fields[2] != "" ? app.parseGrokDate(fields[2]) : "";
                        break;
                    case "SCALAR":
                        row[1] = Number(fields[1]);
                        row[2] = fields[2] != "" ? Number(fields[2]) : "";
                        break;
                    default:
                    case "CATEGORY":
                        row[1] = fields[1];
                        row[2] = fields[2];
                        break;
                }
                row[3] = (typeof fields[3] == "number") ? (fields[3] * 100) + '%' : "";

                if (typeof row[2] == "number") {
                    // http://en.wikipedia.org/wiki/Symmetric_mean_absolute_percentage_error
                    result.SMAPE.n++;
                    result.SMAPE.numerator += (Math.abs(row[2] - row[1]));
                    result.SMAPE.denominator += (row[1] + row[2]);
                }

                result.rows.push(row);
            }

            return result;
        }
    };
    
    global.Sample.ModelMonitor = function(model, opts) {
        this.model = model;
        this._lastStatus = model.get("status");
        this._statusChangeListeners = [];
        GROK.Monitor.call(this, this._modelPoller, opts);
    }

    global.Sample.ModelMonitor.prototype = GROK.util.heir(GROK.Monitor.prototype);
    global.Sample.ModelMonitor.prototype.constructor = GROK.Monitor;

    global.Sample.ModelMonitor.prototype._modelPoller = function (cb) {
        var me = this;
        this._print('polling for model status changes...');

        var project = app.project;
        project.getModel(this.model.getId(), function (err, model) {
            if (err) {
                return me._fire('error', err);
            }

            var status = model.get('status') || GROK.Model.STATUS.STOPPED;
            me._print('model status: ' + status);
            if (status !== me._lastStatus) {
                // for the statusChange listeners
                me.statusChange(model);
            }
            me._lastStatus = status;
            // for the onPoll listeners
            cb(model);
        });
    };

    global.Sample.ModelMonitor.prototype.onStatusChange = function (fn) {
        this._statusChangeListeners.push({ trigger: ANY, fn: fn });
    };

    global.Sample.ModelMonitor.prototype.onStatus = function (status, fn) {
        this._statusChangeListeners.push({ trigger: status, fn: fn });
    };

    global.Sample.ModelMonitor.prototype.statusChange = function (model) {
        this._print('ModelMonitor calling statusChange listeners');
        this._statusChangeListeners.forEach(function (listener) {
            var trigger = listener.trigger,
                fn = listener.fn;
            if (trigger === ANY || trigger === model.get('status')) {
                fn(model);
            }
        });
    };

})(window);