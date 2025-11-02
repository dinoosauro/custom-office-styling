<script lang="ts">
    import { fade } from "svelte/transition";
    import Card from "../lib/Card.svelte";
    import { ExcelCharts, type HelperType } from "../Scripts/HelperType";
    import { lang } from "../Scripts/Language";
    import ChartAxes from "./lib/ChartAxes.svelte";
    import ChartDataLabels from "./lib/ChartDataLabels.svelte";
    import ChartFormat from "./lib/ChartFormat.svelte";
    import ChartLegend from "./lib/ChartLegend.svelte";
    import ChartPlotArea from "./lib/ChartPlotArea.svelte";
    import ChartSeries from "./lib/ChartSeries.svelte";
    import ChartTitle from "./lib/ChartTitle.svelte";
    import { GetChart, UpdateSelectedChart } from "./Scripts/GetChart";
    import UpdateGenericProperties from "./Scripts/UpdateGenericProperties";
    import { cubicInOut } from "svelte/easing";
    import ExportChartImage from "./lib/ExportChartImage.svelte";
    /**
     * The Chart that should be edited
     */
    let data: Excel.Chart;
    /**
     * The ID of the chart that will be edited or exported
     */
    let selectedChartId = "";
    /**
     * Base64 image of the chart. Used only when the user is picking some charts
     */
    let chartImageBase64 = $state("");
    /**
     * Section of the Excel app
     */
    let selectedSection = $state<"chartEdit" | "chartImg" | "none">("none");
    /**
     * Promise used while the app is waiting for user input to choose the chart to edit.
     */
    let waitPromise: (value: string) => void;
    /**
     * A list of all the chart, that contains [its ID, its title, its width, its height]
     */
    let availableCharts: [string, string, number, number][] = $state([]);
    /**
     * Force re-rendering of the chart part
     */
    let forceChart = $state(0);
    /**
     * The section chosen in the "Edit chart" part
     */
    let selectedChart = $state("axis");
    /**
     * Width and height of the output image. Used in the "Chart export" section.
     */
    let outputDimensions = [1920, 1080];

    let { urlCallback }: { urlCallback: (url: string) => void } = $props();
    /**
     * Make the user choose the chart to edit/export
     * @param ctx the Excel context request
     */
    async function pickChart(ctx: Excel.RequestContext) {
        const charts = ctx.workbook.worksheets
            .getActiveWorksheet()
            .charts.load({
                title: { text: true },
                id: true,
                width: true,
                height: true,
            });
        await ctx.sync();
        // Populate the availableCharts array. The picker dialog will be automatically shown.
        availableCharts = charts.items.map((i) => [
            i.id,
            i.title.text || i.name,
            i.width,
            i.height
        ]);
        const id = await new Promise<string>((res) => (waitPromise = res));
        UpdateSelectedChart(id);
        return {
            charts,
            id,
        };
    }
        /**
   * A Spinner element at the center of the page
   */
  const spinner = document.createElement("div");
  spinner.classList.add("spinner");
</script>

{#if selectedSection === "none"}
    <Card>
        <h2>{lang("Edit chart")}:</h2>
        <p>
            {lang("A list of all the available charts will appear. Pick the one you want to edit from the dropdown menu")}.
        </p>
        <button
            onclick={() => {
                document.body.append(spinner);
                setTimeout(() => {
                Excel.run(async (ctx) => {
                    // Ask the user to pick a chart
                    const { charts, id } = await pickChart(ctx);
                    // Let's now load all the items we need to edit. We'll wrap some things in try blocks, so that, if they are not there, the extension will still be usable.
                    let tempData = charts.items.find(
                        (i) => i.id === id,
                    ) as Excel.Chart;
                    let haveSeriesBeenLoaded = true;
                    try {
                        tempData.axes.load({ $all: true });
                        await ctx.sync();
                    } catch(ex) {
                        console.warn(ex, "Failed axis loading");
                    }
                    try {
                        tempData.series.load();
                        await ctx.sync();
                    } catch(ex) {
                        console.warn(ex, "Failed series loading");
                        haveSeriesBeenLoaded = false;
                    }
                    try {
                        tempData.format.load({
                            border: { $all: true },
                            font: { $all: true },
                            $all: true,
                        });
                        await ctx.sync();
                    } catch(ex) {
                        console.warn(ex, "Failed format loading");
                    }
                    try {
                        tempData.plotArea.load({
                            $all: true,
                            format: { border: { $all: true } },
                        });
                        await ctx.sync();
                    } catch (ex) {
                        console.warn(ex);
                    }
                    try {
                        tempData.dataLabels.load({
                            $all: true,
                            format: {
                                $all: true,
                                border: { $all: true },
                                font: { $all: true },
                            },
                        });
                        await ctx.sync();
                    } catch (ex) {
                        console.warn(ex);
                    }
                    try {
                        tempData.title.load({
                            $all: true,
                            format: {
                                $all: true,
                                border: { $all: true },
                                font: { $all: true },
                            },
                        });
                        await ctx.sync();
                    } catch (ex) {
                        console.warn(ex);
                    }
                    if (haveSeriesBeenLoaded) {
                    for (const serie of tempData.series.items) {
                        try {
                            serie.format.load({ $all: true, line: { $all: true } });
                            serie.points.load({
                                $all: true,
                                format: { border: { $all: true } },
                            });
                            await ctx.sync();
                            for (const point of serie.points.items) {
                                point.dataLabel.load({$all: true, format: {font: {$all: true}, border: {$all: true}}})
                                try {
                                    await ctx.sync();
                                } catch(ex) {
                                    console.warn(ex);
                                }
                            }
                        } catch(ex) {
                            console.warn(ex);
                        }
                    }
                    await ctx.sync();
                }
                    try {
                        tempData.legend.load({
                            format: {
                                $all: true,
                                font: { $all: true },
                                border: { $all: true },
                            },
                            $all: true,
                        });
                        await ctx.sync();
                    } catch (ex) {
                        console.warn("No legend found", ex);
                    }
                    for (const prop of [
                        "categoryAxis",
                        "valueAxis",
                        "seriesAxis",
                    ]) {
                        try {
                            tempData.axes[prop as "categoryAxis"].load();
                            tempData.axes[
                                prop as "categoryAxis"
                            ].majorGridlines.format.line.load();
                            tempData.axes[
                                prop as "categoryAxis"
                            ].majorGridlines.load();
                            tempData.axes[
                                prop as "categoryAxis"
                            ].minorGridlines.load();
                            tempData.axes[
                                prop as "categoryAxis"
                            ].minorGridlines.format.line.load();
                            await ctx.sync();
                        } catch (ex) {
                            console.warn(ex);
                        }
                    }
                    data = JSON.parse(JSON.stringify(tempData));
                    selectedSection = "chartEdit";
                    forceChart = Date.now();
                    spinner.remove();
                });
            }, 1)
            }}>{lang("Pick the chart to edit")}</button
        >
    </Card><br />
    <Card>
        <h2>{lang("Export a chart as an image")}</h2>
        <p>
            {lang("Pick a chart, choose its witdh/height, and save it as a PNG image")}.        </p>
        <button
            onclick={() => {
                document.body.append(spinner);
                setTimeout(() => {
                    Excel.run(async (ctx) => {
                        await pickChart(ctx);
                        selectedSection = "chartImg";
                        spinner.remove();
                    });
                }, 1);
            }}>{lang("Pick chart")}</button
        >
    </Card>
{:else if selectedSection === "chartEdit"}
    {#key forceChart}
        {#if forceChart !== 0}
            <br />
            <Card>
                <div class="flex gap" style="overflow: auto">
                    {#each [["axis", lang("Axis options")], ["format", lang("Format options")], ...(data.legend?.visible ? [["legend", lang("Legend options")]] : []), ["plotarea", lang("Chart area options")], ["series", lang("Series options")],["datalabels", lang("Data labels options")], ["title", lang("Title options")], ["other", lang("Other chart options")]] as [key, title]}
                        <button
                            class="card secondCard chip"
                            style={selectedChart === key
                                ? "background-color: var(--accent)"
                                : "background-color: var(--input)"}
                            onclick={() => (selectedChart = key as "list")}
                        >
                            {title}
                        </button>
                    {/each}
                </div>
            </Card><br />
            {#if selectedChart === "axis"}
                <Card>
                    <ChartAxes spinner={spinner} {data}></ChartAxes>
                </Card>
            {:else if selectedChart === "format"}
                <Card>
                    <ChartFormat spinner={spinner} {data}></ChartFormat>
                </Card>
            {:else if selectedChart === "legend"}
                <Card>
                    <ChartLegend spinner={spinner} {data}></ChartLegend>
                </Card>
            {:else if selectedChart === "plotarea"}
                <Card>
                    <ChartPlotArea spinner={spinner} {data}></ChartPlotArea>
                </Card>
            {:else if selectedChart === "series"}
                <Card>
                    <ChartSeries spinner={spinner} {data}></ChartSeries>
                </Card>
            {:else if selectedChart === "datalabels"}
                <Card>
                    <ChartDataLabels spinner={spinner} dataLabels={data.dataLabels}
                    ></ChartDataLabels>
                </Card>
            {:else if selectedChart === "title"}
                <Card>
                    <ChartTitle spinner={spinner} {data}></ChartTitle>
                </Card>
            {:else if selectedChart === "other"}
                <Card>
                    <h2>{lang("Other chart properties")}:</h2>
                    <label class="flex hcenter gap">
                        {lang("Chart type")}:
                        <div class="selectContainer">
                            <select bind:value={data.chartType}>
                                {#each ExcelCharts as chart}
                                    <option value={chart}>{chart}</option>
                                {/each}
                            </select>
                        </div>
                    </label><br />
                    <label class="flex hcenter gap">
                        {lang("Display blank cells as")}:
                        <div class="selectContainer">
                            <select bind:value={data.displayBlanksAs}>
                                <option value="NotPlotted">NotPlotted</option>
                                <option value="Zero">Zero</option>
                                <option value="Interplotted">Interplotted</option>
                            </select>
                        </div>
                    </label><br />
                    <label class="flex hcenter gap">
                        {lang("Use for data")}:
                        <div class="selectContainer">
                            <select bind:value={data.plotBy}>
                                <option value="Rows">{lang("Rows")}</option>
                                <option value="Columns">{lang("Columns")}</option>
                            </select>
                        </div>
                    </label><br />
                    <label class="flex hcenter gap">
                        <input
                            type="checkbox"
                            bind:checked={data.plotVisibleOnly}
                        />{lang("Plot only visible cells")}
                    </label><br />
                    <label class="flex hcenter gap">
                        <input
                            type="checkbox"
                            bind:checked={data.showDataLabelsOverMaximum}
                        />{lang(
                            "Show data labels even if their value is over the maximum displayed in the axis",
                        )}
                    </label><br />
                    <button
                        onclick={() => {
                            document.body.append(spinner);
                            setTimeout(() => {
                                Excel.run(async (ctx) => {
                                    const chart = await GetChart(ctx);
                                    UpdateGenericProperties(data, chart);
                                    await ctx.sync();
                                    spinner.remove();
                                });
                            }, 1)
                        }}>{lang("Apply")}</button
                    >
                </Card>
            {/if}
        {/if}
    {/key}
{:else}
    <Card>
        <ExportChartImage
            {urlCallback}
            height={outputDimensions[1]}
            width={outputDimensions[0]}
            spinner={spinner}
        ></ExportChartImage>
    </Card>
{/if}
{#if availableCharts.length !== 0}
    <div
        class="dialog"
        in:fade={{ duration: 400, easing: cubicInOut }}
        out:fade={{ duration: 400, easing: cubicInOut }}
    >
        <div>
            <h2>{lang("Pick the chart to edit")}:</h2>
            <div class="selectContainer">
                <select
                    onchange={() => {
                        document.body.append(spinner);
                        setTimeout(() => {
                            Excel.run(async (ctx) => {
                                const charts = ctx.workbook.worksheets
                                    .getActiveWorksheet()
                                    .charts.load({ id: true });
                                await ctx.sync();
                                const base64 = charts.items
                                    .find((i) => i.id === selectedChartId)
                                    ?.getImage();
                                await ctx.sync();
                                if (base64?.value) chartImageBase64 = base64.value;
                                spinner.remove()
                            });
                        }, 1)
                    }}
                    bind:value={selectedChartId}
                >
                    {#each availableCharts as [id, name]}
                        <option value={id}>{name}</option>
                    {/each}
                </select>
            </div>
            <br />
            {#if chartImageBase64}
            <img
                alt="Selected chart"
                src={`data:image/png;base64,${chartImageBase64}`}
                style="width: 100%; border-radius: 12px;"
            />
            <br /><br />
            <button
                onclick={() => {
                    waitPromise(selectedChartId);
                    outputDimensions = (availableCharts.find(i => i[0] === selectedChartId)?.slice(2) as [number, number]).map(i => Math.floor(i * 3));
                    availableCharts = [];
                }}>{lang("Pick chart")}</button
            >
            {/if}
        </div>
    </div>
{/if}
