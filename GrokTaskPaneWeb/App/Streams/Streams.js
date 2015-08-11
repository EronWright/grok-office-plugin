/* Common app functionality */

(function (global) {
    "use strict";

    global.Sample.AsyncResult = function () {
        this.asyncContext = null;
        this.status = Office.AsyncResultStatus.Succeeded;
        this.error = null;
        this.value = null;
    }

    global.Sample.StreamBindingFactory = function () {
        this.persistedStreams = Office.context.document.settings.get("Streams") || {};
    }

    global.Sample.StreamBindingFactory.prototype = {

        //getBindingsAsync: function(callback) {
        //    // initialize stream bindings
        //    for (var streamId in this.persistedStreams) {
        //        // find the associated binding
        //        var bindingId = this.persistedStreams[streamId].bindingId;
        //        Office.context.document.bindings.getByIdAsync(bindingId, function (asyncResult) {

        //            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        //                // remove stale binding
        //                delete this.persistedStreams[streamId];
        //            }
        //            else {
        //                var binding = asyncResult.value;
        //            }
        //        });
        //    }
        //},

        // returns a TableBindingToStream instance for the given streamId, or null.

        getBindingForStreamAsync: function (stream, callback) {
            var p = this.persistedStreams[stream.getId()];
            if (!p) {
                var ar = new Sample.AsyncResult();
                ar.value = null;
                callback(ar);
                return;
            }

            var bindingId = p.bindingId;

            Office.context.document.bindings.getByIdAsync(bindingId, $.proxy(function (asyncResult) {

                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error("Table binding is gone.");
                        callback(ar);
                    }
                    return;
                }

                var binding = asyncResult.value;

                var ar = new Sample.AsyncResult();
                ar.value = new Sample.TableBindingToStream(this, stream, binding);
                callback(ar);
            }, this));

        },

        createFromStreamAsync: function (stream, callback) {
            
            // Build a table based on the schema of the stream.
            var table = new Office.TableData();

            var dataSource = stream.getScalar("dataSources")[0];
            table.headers = [$.map(dataSource.fields, function (field) {
                return field.name;
            })];

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
                        { id: stream.getId() },
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

                            this.bindAsync(binding, stream, $.proxy(function (asyncResult) {
                                if (callback) {
                                    var ar = new Sample.AsyncResult();
                                    ar.status = Office.AsyncResultStatus.Succeeded;
                                    ar.value = new Sample.TableBindingToStream(this, stream, binding);
                                    callback(ar);
                                }
                                return;
                            }, this));
                        }, this)
                    );
                },this));
        },

        bindAsync: function (binding, stream, callback) {
            // finalize the binding of a table to a stream
            this.persistedStreams[stream.getId()] = { bindingId: binding.id, nextRow: 0 };
            this.persistAsync(callback);
        },

        releaseBindingAsync: function(binding, streamId, callback) {
            Office.context.document.bindings.releaseByIdAsync(binding.binding.id, $.proxy(function (asyncResult) {
                delete this.persistedStreams[streamId];
                this.persistAsync(callback);
            }, this));
        },

        // infer a stream definition from a table binding
        createDataSourceFromBindingAsync: function (binding, callback) {

            // inspect the first row to infer the schema
            binding.getDataAsync(
                {
                    coercionType: Office.CoercionType.Table, valueFormat: Office.ValueFormat.Formatted,
                    rowCount: 1
                },
                function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        if (callback) {
                            var ar = new Sample.AsyncResult();
                            ar.status = Office.AsyncResultStatus.Failed;
                            ar.error = new Error("Unable to get table data.  Error: " + asyncResult.error.message);
                            callback(ar);
                        }
                        return;
                    }
                
                    var tableData = asyncResult.value;

                    var dataSource = {
                        name: binding.id,
                        dataSourceType: 'local',
                        fields: []
                    };

                    var fieldArray = [];
                    var sampleRow = tableData.rows.length >= 1 ? tableData.rows[0] : null;
                    $.each(tableData.headers[0], function (i, header) {
                        var sampleValue = sampleRow != null ? sampleRow[i] : "";

                        var field = {
                            name: header
                        };
                    
                        if ($.isNumeric(sampleValue)) {
                            field.dataFormat = {
                                dataType: 'SCALAR'
                            };
                        }
                        else if (!isNaN(Date.parse(sampleValue))) {
                            field.flag = 'TIMESTAMP';
                            field.dataFormat = {
                                dataType: 'DATETIME'
                            };
                        }
                        else {
                            field.dataFormat = {
                                dataType: 'CATEGORY'
                            };
                        }

                        fieldArray.push(field);
                    });

                    dataSource.fields = fieldArray;
               
                    var ar = new Sample.AsyncResult();
                    ar.status = Office.AsyncResultStatus.Succeeded;
                    ar.value = dataSource;
                    callback(ar);
                }
            );

        },

        persistAsync: function(callback) {

            Office.context.document.settings.set("Streams", this.persistedStreams);
            Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error("Unable to persist stream bindings.  Error: " + asyncResult.error.message);
                        callback(ar);
                    }
                    return;
                }

                var ar = new Sample.AsyncResult();
                callback(ar);
            });
        }
    }

    global.Sample.TableBindingToStream = function (factory, stream, binding) {
        this.factory = factory;
        this.stream = stream;
        this.binding = binding;
        this.uploadState = { status: "stopped" };
        this._uploadListener = null;
    }

    global.Sample.TableBindingToStream.prototype = {

        setUploadListener: function(listener) {
            this._uploadListener = listener;
        },

        // upload a bound Excel table to a stream
        uploadAsync: function (callback) {
            var startRow = this.factory.persistedStreams[this.stream.getId()].nextRow;

            // refresh the binding to get updated rowCount
            Office.context.document.bindings.getByIdAsync(this.binding.id, $.proxy(function (asyncResult) {

                // note that the binding object is immutable; the rowCount does not change when the user edits the table.
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    if (callback) {
                        var ar = new Sample.AsyncResult();
                        ar.status = Office.AsyncResultStatus.Failed;
                        ar.error = new Error("Table binding is gone.");
                        callback(ar);
                    }
                    return;
                }

                var binding = asyncResult.value;

                this._setUploadState({ status: "running", uploadedCount: 0, rowCount: binding.rowCount });

                this._uploadPageAsync(binding, startRow, $.proxy(function (asyncResult) {

                    this._setUploadState({ status: "stopped" });

                    if (callback) {
                        callback(asyncResult);
                    }
                }, this));
            }, this));

        },

        _uploadPageAsync: function (binding, startRow, callback) {

            if (startRow >= binding.rowCount) {
                // all rows have been uploaded
                console.log("upload complete", binding.rowCount);
                if (callback) {
                    var ar = new Sample.AsyncResult();
                    ar.status = Office.AsyncResultStatus.Succeeded;
                    callback(ar);
                }
                return;
            }
            
            console.log("upload: get table data", startRow);

            binding.getDataAsync(
                {
                    coercionType: Office.CoercionType.Table,
                    valueFormat: Office.ValueFormat.Unformatted,
                    startRow: startRow,
                    rowCount: Math.min(binding.rowCount - startRow, 1000)
                },
                $.proxy(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        console.log("upload: error getting table data", asyncResult.error);
                        if (callback) {
                            var ar = new Sample.AsyncResult();
                            ar.status = Office.AsyncResultStatus.Failed;
                            ar.error = new Error("Unable to get table data.  Error: " + asyncResult.error.message);
                            callback(ar);
                        }
                        return;
                    }

                    var tableData = asyncResult.value;
                    var grokDataSource = this.stream.getScalar("dataSources")[0];

                    console.log("upload: converting table data");
                    var grokData = this.convertTableDataToGrokData(tableData, grokDataSource);

                    // upload page of data
                    console.log("upload: uploading data");
                    this.stream.addData(grokData, $.proxy(function (err) {

                        if (err) {
                            console.log("error uploading to stream", err.message);
                            if (callback) {
                                var ar = new Sample.AsyncResult();
                                ar.status = Office.AsyncResultStatus.Failed;
                                ar.error = new Error(err);
                                callback(ar);
                            }
                            return;
                        }
                        
                        // update startRow
                        startRow = startRow + grokData.length;

                        // persist startRow
                        this.factory.persistedStreams[this.stream.getId()].nextRow = startRow;
                        this.factory.persistAsync($.proxy(function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                                if (callback) {
                                    var ar = new Sample.AsyncResult();
                                    ar.status = Office.AsyncResultStatus.Failed;
                                    ar.error = new Error("Unable to persist stream.  Error: " + asyncResult.error.message);
                                    callback(ar);
                                }
                                return;
                            }

                            // continue
                            this._setUploadState({ status: "running", uploadedCount: startRow, rowCount: binding.rowCount });
                            this._uploadPageAsync(binding, startRow, callback);

                        }, this));
                    }, this));
                }, this)
            );
        },

        _setUploadState: function(obj) {
            this.uploadState = obj;
            if (this._uploadListener) {
                this._uploadListener(this.uploadState);
            }
        },

        convertTableDataToGrokData: function (tableData, dataSource) {

            var grokData = new Array();

            for (var i = 0; i < tableData.rows.length; i++) {
                
                var row = new Array(dataSource.fields.length);
                // use header information 
                for (var field = 0; field < dataSource.fields.length; field++) {

                    var value;
                    switch (dataSource.fields[field].dataFormat.dataType.toUpperCase()) {
                        case "DATETIME":
                            var date = app.parseExcelDate(tableData.rows[i][field]);
                            value = date.toISOString().replace(/T/gi, " ").replace(/Z/gi, "");
                            break;
                        case "SCALAR":
                            value = Number(tableData.rows[i][field]);
                            break;
                        default:
                        case "CATEGORY":
                            value = String(tableData.rows[i][field]);
                            break;
                    }

                    row[field] = value;
                }

                grokData.push(row);
            }


            return grokData;
        },
    };
    
})(window);