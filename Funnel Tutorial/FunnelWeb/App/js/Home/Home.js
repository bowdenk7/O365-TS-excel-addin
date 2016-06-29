define(["require", "exports", '../d3-funnel-charts', '../App'], function (require, exports, d3_funnel_charts_1, App_1) {
    "use strict";
    (function () {
        "use strict";
        var options = {
            bindingID: 'myBinding',
            animationSpeed: 1000
        };
        var binding = null;
        //Sample data
        var sampleHeaders = [['Stage', 'Percent']];
        var sampleRows = [
            ['Applied', 100],
            ['Phone Interview', 80],
            ['On-site Interview', 45],
            ['Given Offer', 30],
            ['Accepted Offer', 12]];
        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                App_1.app.initialize();
                options.animationSpeed = Office.context.document.settings.get('animationSpeed') ? Office.context.document.settings.get('animationSpeed') : 1000;
                $('#sampleButton').click(insertSampleData);
                $('#get-data-from-selection').click(getDataFromSelection);
                $('#animationButton').click(function () {
                    if (options.animationSpeed == 3) {
                        options.animationSpeed = 1000;
                        setAndSave('animationSpeed', 1000);
                    }
                    else {
                        options.animationSpeed = 3;
                        setAndSave('animationSpeed', 3);
                    }
                    if (binding) {
                        displayDataForBinding(binding);
                    }
                });
            });
        };
        //Takes in a string of settingName and string, number, or object of settingValue
        //Creates new corresponding setting, then saves settings to the document
        function setAndSave(settingName, settingValue) {
            if (Office.context.document.settings) {
                Office.context.document.settings.set(settingName, settingValue);
                Office.context.document.settings.saveAsync();
            }
        }
        //Creates TableData of sample data, writes it to selected cell in chart, and binds to it
        function insertSampleData() {
            var sampleData = new Office.TableData(sampleRows, sampleHeaders);
            Office.context.document.setSelectedDataAsync(sampleData, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    App_1.app.showNotification('Could not insert sample data', 'Please choose a different selection range.');
                }
                else {
                    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: options.bindingID }, function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            displayExistingData();
                        }
                        else {
                            App_1.app.showNotification(asyncResult.error.name, asyncResult.error.message);
                        }
                    });
                }
            });
        }
        // Reads data from current document selection and displays a notification
        function getDataFromSelection() {
            Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Table, { id: options.bindingID }, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    displayExistingData();
                }
                else {
                    App_1.app.showNotification(result.error.name, result.error.message);
                }
            });
        }
        //Called once by initialize, plus when new binding is created
        //Simply retrieves the current binding (or defaults) and passes it along to displayDataForBinding
        function displayExistingData() {
            Office.context.document.bindings.getByIdAsync(options.bindingID, function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    binding = result.value;
                    displayDataForBinding(binding);
                    // add data-changed event handler to the binding
                    binding.addHandlerAsync(Office.EventType.BindingDataChanged, function () {
                        displayDataForBinding(binding);
                    });
                }
                else {
                    //Cannot retrieve binding (error or none exists), so pass null binding
                    displayDataForBinding(null);
                }
            });
        }
        //Takes in binding, calls helper function on the binding's data if it's not null, else calls helper on default data
        function displayDataForBinding(binding) {
            if (binding) {
                binding.getDataAsync({ coercionType: Office.CoercionType.Matrix, valueFormat: Office.ValueFormat.Unformatted, filterType: Office.FilterType.OnlyVisible }, function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        displayDataHelper(result.value);
                    }
                    else {
                        App_1.app.showNotification("Error retrieving data from binding.", "Bind to a different range and try again.");
                    }
                });
            }
            else {
                var defaultData = [['Category', 'Number'], ['Clicks', 768], ['Free Downloads', 455], ['Purchases', 211], ['Repeat Purchases', 134]];
                displayDataHelper(defaultData);
            }
        }
        //If data meets requirements, this makes associated FunnelChart, clears container, and draws new chart
        function displayDataHelper(data) {
            if (data.length <= 1 || data[0].length !== 2) {
                App_1.app.showNotification("Improper data", "Please select two columns and at least two rows and try again");
            }
            else {
                var chart = new d3_funnel_charts_1.FunnelChart({
                    data: data,
                    width: 400,
                    height: 250
                });
                $('#container').empty();
                chart.draw('#container', options.animationSpeed);
            }
        }
    })();
});
