define(["require", "exports"], function (require, exports) {
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
