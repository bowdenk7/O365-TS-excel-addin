define("App", ["require", "exports"], function (require, exports) {
    /* Common app functionality */
    "use strict";
    exports.app = {
        initialize: function () {
            $('body').append('<div id="notification-message">' +
                '<div class="padding">' +
                '<div id="notification-message-close"></div>' +
                '<div id="notification-message-header"></div>' +
                '<div id="notification-message-body"></div>' +
                '</div>' +
                '</div>');
            $('#notification-message-close').click(function () {
                $('#notification-message').hide();
            });
        },
        showNotification: function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        }
    };
});
// Configure loading modules from the lib directory,
// except for 'app' ones, which are in a sibling
// directory.
requirejs.config({
    baseUrl: 'js',
    paths: {
        app: '.'
    }
});
// Start loading the main app file. Put all of
// your application logic in there.
requirejs(['bundle.js']);
define("d3-funnel-charts", ["require", "exports"], function (require, exports) {
    "use strict";
    var DEFAULT_HEIGHT = 400, DEFAULT_WIDTH = 600, DEFAULT_BOTTOM_PERCENT = 1 / 3;
    var FunnelChart = (function () {
        function FunnelChart(options) {
            this.data = options.data;
            if (isNaN(this.data[0][1])) {
                this.data = this.data.splice(1, this.data.length);
            }
            this.totalEngagement = 0;
            for (var i = 0; i < this.data.length; i++) {
                this.totalEngagement += this.data[i][1];
            }
            this.width = typeof options.width !== 'undefined' ? options.width : DEFAULT_WIDTH;
            this.height = typeof options.height !== 'undefined' ? options.height : DEFAULT_HEIGHT;
            var bottomPct = typeof options.bottomPct !== 'undefined' ? options.bottomPct : DEFAULT_BOTTOM_PERCENT;
            this._slope = 2 * this.height / (this.width - bottomPct * this.width);
            this._totalArea = (this.width + bottomPct * this.width) * this.height / 2;
        }
        FunnelChart.prototype._getLabel = function (ind) {
            return this.data[ind][0];
        };
        ;
        FunnelChart.prototype._getEngagementCount = function (ind) {
            return this.data[ind][1];
        };
        ;
        FunnelChart.prototype._createPaths = function () {
            /* Returns an array of points that can be passed into d3.svg.line to create a path for the funnel */
            var trapezoids = [];
            function findNextPoints(chart, prevLeftX, prevRightX, prevHeight, dataInd) {
                // reached end of funnel
                if (dataInd >= chart.data.length)
                    return;
                // math to calculate coordinates of the next base
                var area = chart.data[dataInd][1] * chart._totalArea / chart.totalEngagement;
                var prevBaseLength = prevRightX - prevLeftX;
                var nextBaseLength = Math.sqrt((chart._slope * prevBaseLength * prevBaseLength - 4 * area) / chart._slope);
                var nextLeftX = (prevBaseLength - nextBaseLength) / 2 + prevLeftX;
                var nextRightX = prevRightX - (prevBaseLength - nextBaseLength) / 2;
                var nextHeight = chart._slope * (prevBaseLength - nextBaseLength) / 2 + prevHeight;
                var points = [[nextRightX, nextHeight]];
                points.push([prevRightX, prevHeight]);
                points.push([prevLeftX, prevHeight]);
                points.push([nextLeftX, nextHeight]);
                points.push([nextRightX, nextHeight]);
                trapezoids.push(points);
                findNextPoints(chart, nextLeftX, nextRightX, nextHeight, dataInd + 1);
            }
            findNextPoints(this, 0, this.width, 0, 0);
            return trapezoids;
        };
        FunnelChart.prototype.draw = function (elem, speed) {
            var DEFAULT_SPEED = 2.5;
            speed = typeof speed !== 'undefined' ? speed : DEFAULT_SPEED;
            var funnelSvg = d3.select(elem).append('svg:svg')
                .attr('width', this.width)
                .attr('height', this.height)
                .append('svg:g');
            // Creates the correct d3 line for the funnel
            var funnelPath = d3.svg.line()
                .x(function (d) { return d[0]; })
                .y(function (d) { return d[1]; });
            // Automatically generates colors for each trapezoid in funnel
            var colorScale = d3.scale.category10();
            var paths = this._createPaths();
            function drawTrapezoids(funnel, i) {
                var trapezoid = funnelSvg
                    .append('svg:path')
                    .attr('d', function (d) {
                    return funnelPath([paths[i][0], paths[i][1], paths[i][2],
                        paths[i][2], paths[i][1], paths[i][2]]);
                })
                    .attr('fill', '#fff');
                var nextHeight = paths[i][paths[i].length - 1];
                var node = trapezoid.node();
                var totalLength = node.getTotalLength();
                var transition = trapezoid
                    .transition()
                    .duration(totalLength / speed)
                    .ease("linear")
                    .attr("d", function (d) { return funnelPath(paths[i]); })
                    .attr("fill", function (d) { return colorScale('#fff'); });
                funnelSvg
                    .append('svg:text')
                    .text(funnel._getLabel(i) + ': ' + funnel._getEngagementCount(i))
                    .attr("x", function (d) { return funnel.width / 2; })
                    .attr("y", function (d) {
                    return (paths[i][0][1] + paths[i][1][1]) / 2;
                }) // Average height of bases
                    .attr("text-anchor", "middle")
                    .attr("dominant-baseline", "middle")
                    .attr("fill", "#fff");
                if (i < paths.length - 1) {
                    transition.each('end', function () {
                        drawTrapezoids(funnel, i + 1);
                    });
                }
            }
            drawTrapezoids(this, 0);
        };
        ;
        return FunnelChart;
    }());
    exports.FunnelChart = FunnelChart;
});
define("Home/Home", ["require", "exports", "d3-funnel-charts", "App"], function (require, exports, d3_funnel_charts_ts_1, App_ts_1) {
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
                App_ts_1.app.initialize();
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
                    App_ts_1.app.showNotification('Could not insert sample data', 'Please choose a different selection range.');
                }
                else {
                    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: options.bindingID }, function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            displayExistingData();
                        }
                        else {
                            App_ts_1.app.showNotification(asyncResult.error.name, asyncResult.error.message);
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
                    App_ts_1.app.showNotification(result.error.name, result.error.message);
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
                        App_ts_1.app.showNotification("Error retrieving data from binding.", "Bind to a different range and try again.");
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
                App_ts_1.app.showNotification("Improper data", "Please select two columns and at least two rows and try again");
            }
            else {
                var chart = new d3_funnel_charts_ts_1.FunnelChart({
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
