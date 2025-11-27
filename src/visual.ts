import * as d3 from "d3";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import DataView = powerbi.DataView;

export class Visual implements IVisual {
    private svg: d3.Selection<SVGSVGElement, unknown, HTMLElement, any>;

    // Constants
    private margin = { left: 100, right: 30, top: 60, bottom: 30 };
    private circleRadius = 14;
    private strokeWidth = 2.5;
    private colourPalette = [
        '#0079EB','#375071','#A83DA4','#4289B3','#9B7003','#293241','#006560','#4B195D'
    ];

    constructor(options: VisualConstructorOptions) {
        this.svg = d3.select(options.element).append("svg");
    }

    public update(options: VisualUpdateOptions) {
        const width = options.viewport.width;
        const height = options.viewport.height;

        this.svg
            .attr("width", width)
            .attr("height", height)
            .selectAll("*")
            .remove(); // clear previous render

        const dataView: DataView = options.dataViews[0];
        if (!dataView || !dataView.categorical || !dataView.categorical.categories || !dataView.categorical.values || dataView.categorical.values.length < 2) {
            return;
        }

        // Extract dynamic data
        const categories = dataView.categorical.categories[0].values;
        const value1 = dataView.categorical.values[0].values;
        const value2 = dataView.categorical.values[1].values;

        const data = categories.map((c, i) => ({
            Categories: String(c),
            Value1: Number(value1[i]),
            Value2: Number(value2[i])
        }));

        // Scales
        const xScale = d3.scaleLinear()
            .domain([0, d3.max(data, d => Math.max(d.Value1, d.Value2))!])
            .range([this.margin.left, width - this.margin.right]);

        const yScale = d3.scalePoint()
            .domain(data.map(d => d.Categories))
            .range([this.margin.top, height - this.margin.bottom]);

        const colourScale = d3.scaleOrdinal()
            .domain(data.map(d => d.Categories))
            .range(this.colourPalette);

        // Horizontal Gridlines
        this.svg.selectAll(".horizontalGridlines")
            .data(data)
            .enter()
            .append("line")
            .attr("class", "horizontalGridlines")
            .attr("x1", this.margin.left)
            .attr("x2", d => xScale(d.Value1) - this.circleRadius)
            .attr("y1", d => yScale(d.Categories)!)
            .attr("y2", d => yScale(d.Categories)!)
            .attr("opacity", 0.9)
            .attr("stroke-width", 1)
            .attr("stroke-dasharray", "3 3")
            .attr("stroke", "#CACACA");

        // Lines between Value1 and Value2
        this.svg.selectAll(".linesV1V2")
            .data(data)
            .enter()
            .append("line")
            .attr("class", "linesV1V2")
            .attr("x1", d => xScale(d.Value1) + this.circleRadius)
            .attr("x2", d => xScale(d.Value2) - this.circleRadius)
            .attr("y1", d => yScale(d.Categories)!)
            .attr("y2", d => yScale(d.Categories)!)
            .attr("stroke-width", this.strokeWidth)
            .attr("stroke", d => colourScale(d.Categories)!);

        // Value1 circles
        this.svg.selectAll(".circles1")
            .data(data)
            .enter()
            .append("circle")
            .attr("class", "circles1")
            .attr("cx", d => xScale(d.Value1))
            .attr("cy", d => yScale(d.Categories)!)
            .attr("r", this.circleRadius)
            .attr("fill", "white")
            .attr("stroke-width", this.strokeWidth)
            .attr("stroke", d => colourScale(d.Categories)!);

        // Value2 circles
        this.svg.selectAll(".circles2")
            .data(data)
            .enter()
            .append("circle")
            .attr("class", "circles2")
            .attr("cx", d => xScale(d.Value2))
            .attr("cy", d => yScale(d.Categories)!)
            .attr("r", this.circleRadius)
            .attr("fill", d => colourScale(d.Categories)!)
            .attr("stroke-width", this.strokeWidth)
            .attr("stroke", d => colourScale(d.Categories)!);

        // Value1 Labels
        this.svg.selectAll(".value1Labels")
            .data(data)
            .enter()
            .append("text")
            .attr("class", "value1Labels")
            .attr("x", d => xScale(d.Value1))
            .attr("y", d => yScale(d.Categories)!)
            .attr("text-anchor", "middle")
            .attr("alignment-baseline", "middle")
            .attr("fill", d => colourScale(d.Categories)!)
            .style("font-size", `${this.circleRadius - 2}px`)
            .style("font-weight", "bold")
            .text(d => d.Value1);

        // Value2 Labels
        this.svg.selectAll(".value2Labels")
            .data(data)
            .enter()
            .append("text")
            .attr("class", "value2Labels")
            .attr("x", d => xScale(d.Value2))
            .attr("y", d => yScale(d.Categories)!)
            .attr("text-anchor", "middle")
            .attr("alignment-baseline", "middle")
            .attr("fill", "white")
            .style("font-size", `${this.circleRadius - 2}px`)
            .style("font-weight", "bold")
            .text(d => d.Value2);

        // Optional: Category Labels with wrap
        this.svg.selectAll(".categoryLabels")
            .data(data)
            .enter()
            .append("text")
            .attr("class", "categoryLabels")
            .attr("x", this.margin.left - 10)
            .attr("y", d => yScale(d.Categories)!)
            .attr("text-anchor", "end")
            .attr("alignment-baseline", "middle")
            .text(d => d.Categories)
            .call(this.wrap, 100); // 100px max width
    }

    // Text wrapping function
    private wrap(textSelection: d3.Selection<d3.BaseType, any, any, any>, width: number) {
        textSelection.each(function () {
            const text = d3.select(this);
            const words = text.text().split(/\s+/).reverse();
            let word: string | undefined;
            let line: string[] = [];
            let lineNumber = 0;
            const lineHeight = 1.2; // ems

            text.text(null);

            let tspan = text.append("tspan")
                .attr("x", 0)
                .attr("dy", "0em");

            while ((word = words.pop())) {
                line.push(word);
                tspan.text(line.join(" "));
                if (tspan.node()!.getComputedTextLength() > width) {
                    line.pop();
                    tspan.text(line.join(" "));
                    line = [word];
                    tspan = text.append("tspan")
                        .attr("x", 0)
                        .attr("dy", `${lineHeight}em`)
                        .text(word);
                }
            }
        });
    }
}
