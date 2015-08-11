/// <reference path="../App.js" />

(function () {
    "use strict";

    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize({
                noproject: true,
            },
            $.proxy(function() {
                
                $('#project-create-button').button({
                }).click(_onCreateClicked);

                var projectList = $('#project-list')[0];
                app.client.listProjects(function (err, projects) {

                    if (projects.length < 1) {
                        $('#projects-panel').hide();
                        return;
                    }

                    $.each(projects, function (i, project) {
                        projectList[projectList.length] = new Option(project.getName(), project.getId());
                    });

                    $('#project-select-button').click(function () { _onProjectSelected({ id: $('#project-list').val() }); });
                });

            }, this));
        });
    };

    function _onCreateClicked() {

        var projectDef = {
            name: $('#project-create-name').val(),
            description: $('#project-create-description').val(),
        };

        if (!_validate(projectDef)) return;

        app.client.createProject(projectDef, function (err, project) {
            if (err) {
                app.showNotification('Create Project', err);
                return false;
            }

            var projectRef = { id: project.getId() };

            _onProjectSelected(projectRef);
            
        });
    };

    function _validate(projectDef) {
        if (projectDef.name.length < 1) return false;

        return true;
    }

    function _onProjectSelected(projectRef) {
        // persist the project selection into the document, so that the page isn't seen again.
        Office.context.document.settings.set("Project", projectRef);
        Office.context.document.settings.saveAsync({ overwriteIfStale: true }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification("Error", "Unable to save the project selection.  Error: " + asyncResult.error.message);
                return;
            }

            window.navigate("/App/Home/Home.html");
        });
    }
})();